#coding=utf-8
from openpyxl import load_workbook
from transliterate import translit
import pymorphy2
morph = pymorphy2.MorphAnalyzer()

wb = load_workbook(filename = 'Hockey_Log.xlsx')
ws = wb.get_sheet_by_name(wb.get_sheet_names()[1])
lang = 'ru'

def get_lang(s):
	en = 'abcdefghijklmnopqrstuwvxyz'
	for let in en:
		if let in s.lower():
			return 'en'
	return 'ru'

def l(str):
	if get_lang(str) != lang:
		return translit(str, lang).encode('utf-8')
	return str

def find_row(ws, ttl):
	idx = 0
	for i in ws.columns[0]:
		idx = idx + 1
		if str(i.value).lower().find(ttl.lower()) > -1:
			return idx
	return -1

def find_empty_row(ws, r):
	for i in range(r, len(ws.rows)-1):
		e = True
		for c in ws.rows[i]:
			if c.value != None:
				e = False
		if e:
			return i
	return len(ws.rows)

def find_team(ws, ttl):
	start = find_row(ws, ttl)
	end = find_empty_row(ws, start)
	out = {}
	for i in range(start + 1, end):
		out[ws.rows[i][1].value] = [ws.rows[i][3].value, ws.rows[i][2].value]
	return out

def makelog(ws, start, finish, t1, t2):
	out = []
	for row in range(start, finish):
		r = ws.rows[row]
		team = "1" if r[3].value == t1 else "2"
		#			minute      team   event      result 	   who         from        to          object
		out.append([r[1].value, team, r[2].value, r[9].value, r[4].value, r[5].value, r[6].value, r[7].value])
	return out


def get_log(ws, t1, t2):
	out = {}
	start = find_row(ws, 'Play')
	end = find_row(ws, 'End of first')
	out[0] = makelog(ws, start, end - 1, t1, t2)
	start = end
	end = find_row(ws, 'End of second')
	out[1] = makelog(ws, start, end - 1, t1, t2)
	start = end
	end = find_row(ws, 'End of third')
	out[2] = makelog(ws, start, end - 1, t1, t2)
	start = end
	end = find_row(ws, 'End of overtime')
	out[3] = []
	if end > -1:
		out[3] = makelog(ws, start, end - 1, t1, t2)
	start = find_row(ws, 'Shootout')
	end = find_row(ws, 'Final')
	out[4] = []
	if end > -1:
		for row in range(start, end):
			r = ws.rows[row]
			team = "1" if r[3].value == t1 else "2"
			out[4].append([team, r[4].value, r[9].value])
	return out

def find_player(tm, n, t1, t2):
	team = t2
	if str(tm) == '1':
		team = t1
	#for key, player in enumerate(team):
	#	if key == n:
	#		return player[0]
	return team[n]
	return ""

def describe_time(t_log, t1, t2, name1, name2, time_name):
	if len(t_log) < 2:
		return ""
	flect = [morph.parse(tn)[0] for tn in time_name.split(' ')]
	out = "В"
	if time_name[0] == u'в':
		out = out + "o"
	out = out +" " + " ".join([word.inflect({'sing', 'loct'}).word.encode('utf-8') for word in flect])
	s1 = 0
	s2 = 0
	for event in t_log:
		if event[3] == 'scored' or event[2] == 'score':
			if event[1] == '1':
				s1 = s1 + 1
			else:
				s2 = s2 + 1
	if s1 == s2:
		out = out + " количество шайб совпало. \n"
	elif max(s2, s1) > 1:
		winner = name1 if s1 > s2 else name2
		ending = 'у' if max(s1, s2) < 2 else '' if  max(s1, s2) > 5 else 'ы'
		out = out + " \""+winner+"\" забил " + str(max(s1, s2)) + " шайб" + ending + ".\n"
	else:

	return out

def describe_log(log, t1, t2, name1, name2, s1, s2):
	out = ""
	#out = out + l(name2) + "\n"
	score_adj = ''
	verb = "победил"
	n2 = name2
	if s1 == 0 or s2 == 0:
		score_adj = " сухим "
	elif abs(s1 - s2) > 5:
		score_adj = " разгромным "
	flect = morph.parse(name2.decode('utf-8'))[0]
	if s2 < s1:
		verb = "проиграл"
		n2 = flect.inflect({'sing', 'datv'}).word.capitalize().encode('utf-8')
	elif s2 == s1:
		verb = "сыграл вничью с"
		n2 = flect.inflect({'sing', 'ablt'}).word.capitalize().encode('utf-8')
	#print n2
	out = out + "«" + name1.capitalize() + "» со " + score_adj+ " счётом "+str(s1)+":"+str(s2)+" " +verb+ " «" +n2+ "» в матче Континентальной хоккейной лиги (КХЛ). \n\n"
	
	for i in range(0, 4):
		times = [u'первый тайм', u'второй тайм', u'третий тайм', u'овертайм']
		out = out + describe_time(log[i], t1, t2, name1, name2, times[i])
	#out = out + "%team1_title% забросили %team2_title_dat% 12 %score=0:безответных% шайб."
	# шаблон - равная игра, неравная; равная: результативная и нерезультативная
	# Уже на шестой минуте главный тренер минчан Любомир Покович взял тайм-аут – Дмитрий Мильчаков к этому времени дважды вынул шайбу из сетки после бросков Мозеса и Мяки. 
	# Учитывая, что двух голов финнам в предыдущих победных матчах хватало для достижения успеха, можно заявлять, что Покович с тайм-аутом запоздал. 
	#Однако вряд ли кто-то мог подумать, что по-настоящему «Йокерит» еще не разыгрался. Третью шайбу на последней минуте периода провел Саллинен. 

	# Во втором отрезке «шуты» отличились еще трижды, причем эти голы состоялись в последние три минуты перед перерывом. 
	# Сначала забил Мозес, затем дважды в большинстве ворота белорусов поразили Аалтонен и Хухтала. 
	# Вряд ли при этом можно утверждать, что «Динамо» играло совсем скверно – это не так. 
	# Моменты гости время от времени создавали, в скорости не уступали. Но, тем не менее, шайбы залетали только в их ворота. Хенрик Карлссон, между тем, продолжал увеличивать продолжительность своей «сухой» серии – последний гол от минчан он пропустил во втором матче. 

	#Ждать чудес в заключительном периоде не стоило, и команды, хоть и не отбывали номер, провели его аккуратно и спокойно. Седьмую шайю забросил Коукал, а вскоре после этого Шарль Лингле наконец-то пробил Карлссона, который с досады подкинул вверх свою клюшку – швед не пропускал на протяжении 151 минуты 27 секунд.
	# генерация острых моментов по убыванию

	bullets = -1
	out_b = ""
	for bullet in log[4]:
		if bullets < 0:
			bullets = 0
		if bullet[2] == 'scored':
			bullets = bullets + 1
			p = find_player(bullet[0], bullet[1], t1, t2)
			out_b = out_b + p[0].capitalize().encode('utf-8') + " " + p[1].encode('utf-8') + " забил в серии буллитов.\n"
			#print find_player(bullet[0], bullet[1], t1, t2)
	if bullets < 0:
		out = out + "Серия буллитов не принесла очков ни одной из сторон. \n"
	elif bullets > 2:
		bs1 = 0
		bs2 = 0
		for bullet in log[4]:
			if bullet[2] == 'scored':
				if bullet[0] == 1:
					bs1 = bs1 + 1
				else:
					bs2 = bs2 + 1
		if bs1 > 0:
			bs1 = " " + str(bs1) + " очков команде «" + name1 + "»"
		else:
			bs1 = ''
		if bs2 > 0:
			bs2 = " " + str(bs2) + " очков команде «" + name2 + "»"
			if len(bs1) > 0:
				bs2 = " и" + bs2
		else:
			bs2 = ''
		out = out + "Серия буллитов принесла"+bs1+bs2+". \n"
	else:
		out = out + out_b
	winner = name1 if s1 > s2 else name2
	if abs(s1 - s2) > 2 and abs(s1 - s2) < 7:
		out = out + winner + " уверенно и с запасом обыграл соседа по турнирной таблице регулярного чемпионата, показал цельную и сбалансированную игру и определенно должен вызывать беспокойство у любого соперника, с которым может встретиться в следующем раунде.\n"
	elif s1 == s2:
		out = out + "Матч закончился вничью.\n"
	else:
		out = out + "Матч получился просто разгромным.\n"
	return out

if __name__ == '__main__':
	with open('report_'+lang+".txt", 'w') as fh:
		team = find_team(ws, 'home-team')
		team2 = find_team(ws, 'guest-team')
		t_name = ws.rows[1][4].value.encode('utf-8')
		t_name2 = ws.rows[1][5].value.encode('utf-8')
		scoreline = find_row(ws, 'Final score')
		score1 = ws.rows[scoreline][1].value
		score2 = ws.rows[scoreline+1][1].value
		log = get_log(ws, t_name, t_name2)
		out = describe_log(log, team, team2, t_name, t_name2, score1, score2)
		#print t_name2, ', ', score1, score2
		fh.write(out)