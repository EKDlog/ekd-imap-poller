#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EKD Logistics — Poller IMAP -> Webhook n8n
------------------------------------------
Relève les nouveaux e-mails de la boîte info@ekdlogistics.eu (o2switch) et les
transmet, un par un, au webhook n8n qui déclenche la chaîne de détection
(dédup, rattachement prospect, classification, statut, alerte).

Caractéristiques importantes :
- Connexion IMAP en LECTURE SEULE (EXAMINE) + BODY.PEEK : ne marque JAMAIS
  les e-mails comme lus, ne déplace/supprime rien.
- Détection par UID (identifiant serveur), PAS par lu/non-lu : insensible au
  fait que vous consultiez la boîte dans Outlook.
- Suivi du dernier UID traité dans un fichier d'état (state.json) : pas de
  double traitement. n8n dédoublonne en plus par Message-ID.
- Aucune dépendance externe : uniquement la bibliothèque standard Python 3.

Configuration : via variables d'environnement, sinon via un fichier .env
placé à côté de ce script (format KEY=VALUE, une par ligne).
Clés : IMAP_HOST, IMAP_PORT, IMAP_USER, IMAP_PASSWORD,
       N8N_WEBHOOK_URL, N8N_WEBHOOK_SECRET,
       MAILBOX (def. INBOX), STATE_FILE (def. state.json),
       LOOKBACK_DAYS (def. 3, utilisé seulement au 1er lancement).
"""
import os, sys, json, ssl, imaplib, email, urllib.request, urllib.error
from email.header import decode_header, make_header
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))

def load_config():
    cfg = {}
    envfile = os.path.join(HERE, '.env')
    if os.path.exists(envfile):
        with open(envfile, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                k, v = line.split('=', 1)
                cfg[k.strip()] = v.strip().strip('"').strip("'")
    # OS env overrides .env
    for k in ('IMAP_HOST','IMAP_PORT','IMAP_USER','IMAP_PASSWORD','N8N_WEBHOOK_URL',
              'N8N_WEBHOOK_SECRET','MAILBOX','STATE_FILE','LOOKBACK_DAYS'):
        if os.environ.get(k):
            cfg[k] = os.environ[k]
    cfg.setdefault('IMAP_PORT', '993')
    cfg.setdefault('MAILBOX', 'INBOX')
    cfg.setdefault('STATE_FILE', os.path.join(HERE, 'state.json'))
    cfg.setdefault('LOOKBACK_DAYS', '3')
    return cfg

def log(*a):
    print('[EKD poller]', datetime.utcnow().isoformat(timespec='seconds'), *a, flush=True)

def dec(s):
    if s is None:
        return ''
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return str(s)

def get_bodies(msg):
    text_plain, text_html = '', ''
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get('Content-Disposition') or '')
            if 'attachment' in disp.lower():
                continue
            if ctype == 'text/plain' and not text_plain:
                text_plain = decode_part(part)
            elif ctype == 'text/html' and not text_html:
                text_html = decode_part(part)
    else:
        ctype = msg.get_content_type()
        payload = decode_part(msg)
        if ctype == 'text/html':
            text_html = payload
        else:
            text_plain = payload
    return text_plain, text_html

def decode_part(part):
    try:
        raw = part.get_payload(decode=True)
        if raw is None:
            return ''
        charset = part.get_content_charset() or 'utf-8'
        return raw.decode(charset, errors='replace')
    except Exception:
        try:
            return str(part.get_payload())
        except Exception:
            return ''

def load_state(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_state(path, state):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(state, f)
    os.replace(tmp, path)

def post_webhook(url, payload, timeout=30):
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, method='POST',
                                 headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.status

WANTED_HEADERS = ['auto-submitted','precedence','x-autoreply','x-autorespond',
                  'x-auto-response-suppress','message-id','from','subject','date',
                  'return-path','x-failed-recipients','content-type']

def main():
    cfg = load_config()
    for req in ('IMAP_HOST','IMAP_USER','IMAP_PASSWORD','N8N_WEBHOOK_URL','N8N_WEBHOOK_SECRET'):
        if not cfg.get(req):
            log('ERREUR: configuration manquante:', req)
            sys.exit(2)

    state = load_state(cfg['STATE_FILE'])
    last_uid = int(state.get('last_uid', 0) or 0)

    ctx = ssl.create_default_context()
    imap = imaplib.IMAP4_SSL(cfg['IMAP_HOST'], int(cfg['IMAP_PORT']), ssl_context=ctx)
    try:
        imap.login(cfg['IMAP_USER'], cfg['IMAP_PASSWORD'])
        # LECTURE SEULE : EXAMINE au lieu de SELECT -> aucun flag modifié
        imap.select(cfg['MAILBOX'], readonly=True)

        if last_uid > 0:
            crit = ['UID', '%d:*' % (last_uid + 1)]
        else:
            since = (datetime.utcnow() - timedelta(days=int(cfg['LOOKBACK_DAYS']))).strftime('%d-%b-%Y')
            crit = ['SINCE', since]
        typ, data = imap.uid('SEARCH', None, *crit)
        if typ != 'OK':
            log('SEARCH KO:', typ, data); return
        uids = [int(x) for x in (data[0].split() if data and data[0] else [])]
        # UID n:* renvoie toujours au moins le dernier message : filtrer <= last_uid
        uids = sorted(u for u in uids if u > last_uid)
        log('UIDs à traiter:', uids if uids else '(aucun)')

        sent = 0
        for uid in uids:
            typ, msgdata = imap.uid('FETCH', str(uid), '(BODY.PEEK[])')
            if typ != 'OK' or not msgdata or not msgdata[0]:
                log('FETCH KO uid', uid); break
            raw = msgdata[0][1]
            msg = email.message_from_bytes(raw)
            text_plain, text_html = get_bodies(msg)
            headers = {}
            for h in WANTED_HEADERS:
                v = msg.get(h)
                if v is not None:
                    headers[h] = dec(v) if h in ('from','subject') else str(v)
            payload = {
                'secret': cfg['N8N_WEBHOOK_SECRET'],
                'from': dec(msg.get('From')),
                'subject': dec(msg.get('Subject')),
                'textPlain': text_plain,
                'textHtml': text_html,
                'messageId': (msg.get('Message-ID') or '').strip(),
                'date': str(msg.get('Date') or ''),
                'headers': headers,
                'uid': uid,
            }
            try:
                status = post_webhook(cfg['N8N_WEBHOOK_URL'], payload)
                if 200 <= status < 300:
                    sent += 1
                    last_uid = uid
                    state['last_uid'] = last_uid
                    save_state(cfg['STATE_FILE'], state)
                    log('OK uid', uid, 'from', payload['from'][:60], '->', status)
                else:
                    log('Webhook non-2xx', status, 'uid', uid, '- arrêt, reprise au prochain run')
                    break
            except urllib.error.URLError as e:
                log('Webhook erreur réseau uid', uid, ':', e, '- arrêt, reprise au prochain run')
                break
        log('Terminé. Emails transmis:', sent, '| dernier UID:', last_uid)
    finally:
        try:
            imap.logout()
        except Exception:
            pass

if __name__ == '__main__':
    main()
