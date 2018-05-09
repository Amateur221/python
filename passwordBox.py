#!python2

PASSWORDS={'email':'1204377673@qq.com',
			'qq':'1204377673'}

import sys,pyperclip
if len(sys.argv)<2:
	print('Usage: py PASSWORDS.py[account] -copy account password')
	sys.exit()

account=sys.argv[1]
if account in PASSWORDS:
	pyperclip.copy(PASSWORDS[account])
	print('this password  copied to clipboard')
else:
	print('this password is not exist')