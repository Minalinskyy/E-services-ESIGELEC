file = open('/home/e_service/edt_notes_abs.log','r')
a = file.read().split('-'*20)
if len(a) > 9:# take 8
	b = ''
	for i in range(len(a)-9, len(a)-1):
		b += a[i]
		b += '-'*20
	b += a[-1]
	print(b)
	file.close()
	file = open('/home/e_service/abs.log','w')
	file.write(b)

file.close()
