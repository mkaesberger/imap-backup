import argparse
import os, re, binascii
import imaplib, mailbox, email
from tqdm import tqdm


def b64(x):
    x = binascii.b2a_base64(x.encode('utf-16be'))

    x = x.replace(b'/', b',')
    x = x.replace(b'\n', b'')
    x = x.replace(b'=', b'')

    return '&{0}-'.format(x.decode())

def ub64(x):
    x = x.replace(',', '/') + '==='
    x = binascii.a2b_base64(x)
    x = x.decode('utf-16be')

    return x

# modified utf-7 encoding
def encode(x):
    res, off = '', 0

    while off < len(x):
        if x[off] == '&':
            # encode ampersand
            res += '&-'
            off += 1
        else:
            end = 0

            # scan for end of non-ascci
            while (off + end) < len(x) and (ord(x[off + end]) < 0x20 or ord(x[off + end]) > 0x7e):
                end += 1

            if end > 0:
                # non-ascii char sequence
                res += b64(x[off:off + end])
            else:
                # ascii char
                res += x[off]

            off += max(0, end - 1) + 1

    return res

# modified utf-7 encoding
def decode(x):
    res, off = '', 0

    while off < len(x):
        # ascii char
        if x[off] != '&':
            res += x[off]
            off += 1
        else:
            end = 0

            # scan for end
            while (off + end) < len(x) and x[off + end] !=  '-':
                end += 1

            if end == 1:
                # ampersand encoded with zero-length
                res += '&'
            else:
                # non-ascii char sequence
                res += ub64(x[off + 1:off + end])

            off += end + 1

    return res

def main(host, port, username, password, use_ssl=True):
    imap_cls = imaplib.IMAP4_SSL if use_ssl else imaplib.IMAP4
    destination = username

    with imap_cls(host=host, port=port) as imap:
        # login and preparations
        resp, _ = imap.login(username, password)
        if resp.lower() != 'ok':
            raise Exception('authentication failed')

        if not os.path.isdir(destination):
            os.mkdir(destination)

        # list folders
        resp, folders = imap.list()
        assert(resp.lower() == 'ok')

        folder_re = re.compile(r'\((?P<flags>[\S ]*)\) "(?P<delim>[\S ]+)" (?P<name>.+)')

        for item in folders:
            # decode list response
            match = re.search(folder_re, item.decode())
            name = match.group('name')

            if name.startswith('"') and name.endswith('"'):
                name = name[1:-1]

            folder_encoded = name
            folder_decoded = decode(name)

            # select folder
            resp, _ = imap.select('"{0}"'.format(folder_encoded))
            assert(resp.lower() == 'ok')

            # create mbox for folder
            mbox = mailbox.mbox(os.path.join(destination, '{0}.mbox'.format(folder_decoded)), create=True)
            try:
                # discover emails
                resp, mail_list = imap.uid('search', None, 'ALL')
                assert(resp.lower() == 'ok')

                # download and store them in the mbox
                uid_list = tqdm(mail_list[0].split(),
                        bar_format='{l_bar}{bar}|',
                        desc=folder_decoded.rjust(24), miniters=1)
                for uid in uid_list:
                    resp, data = imap.uid('fetch', uid, '(RFC822)')
                    assert(resp.lower() == 'ok')

                    mail = data[0][1]
                    mail = email.message_from_bytes(mail)

                    mbox.add(mail)
            finally:
                # close mbox
                mbox.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--host', type=str, required=True)
    parser.add_argument('--port', type=int, required=True)
    parser.add_argument('--ssl', action='store_true')
    parser.add_argument('--username', type=str, required=True)
    parser.add_argument('--password', type=str, required=True)

    args = parser.parse_args()

    main(args.host, args.port, args.username, args.password, use_ssl=args.ssl)
