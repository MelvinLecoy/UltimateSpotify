from copy import *
import json,docx,requests,base64
from urllib.parse import urlencode
doc=docx.Document()
clientid='a53631fb8367429bb481dc9ebc32be88'
clientsecret='a96a1cd92e2c4fd39f563ef5afa70e15'


class Miscellaneous:
    def readjson(self):
        global datas
        if nof==1:
            with open('StreamingHistory0.json') as a: datas=json.load(a)
        elif nof==2:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            datas=fir+sec
        elif nof==3:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            datas=fir+sec+thir
        elif nof==4:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            datas=fir+sec+thir+four
        elif nof==5:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            with open('StreamingHistory4.json') as e: five=json.load(e)
            datas=fir+sec+thir+four+five
        elif nof==6:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            with open('StreamingHistory4.json') as e: five=json.load(e)
            with open('StreamingHistory5.json') as f: six=json.load(f)
            datas=fir+sec+thir+four+five+six
        elif nof==7:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            with open('StreamingHistory4.json') as e: five=json.load(e)
            with open('StreamingHistory5.json') as f: six=json.load(f)
            with open('StreamingHistory6.json') as g: seven=json.load(g)
            datas=fir+sec+thir+four+five+six+seven
        elif nof==8:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            with open('StreamingHistory4.json') as e: five=json.load(e)
            with open('StreamingHistory5.json') as f: six=json.load(f)
            with open('StreamingHistory6.json') as g: seven=json.load(g)
            with open('StreamingHistory7.json') as h: eight=json.load(h)
            datas=fir+sec+thir+four+five+six+seven+eight
        elif nof==9:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            with open('StreamingHistory4.json') as e: five=json.load(e)
            with open('StreamingHistory5.json') as f: six=json.load(f)
            with open('StreamingHistory6.json') as g: seven=json.load(g)
            with open('StreamingHistory7.json') as h: eight=json.load(h)
            with open('StreamingHistory8.json') as i: nine=json.load(i)
            datas=fir+sec+thir+four+five+six+seven+eight+nine
        elif nof==10:
            with open('StreamingHistory0.json') as a: fir=json.load(a)
            with open('StreamingHistory1.json') as b: sec=json.load(b)
            with open('StreamingHistory2.json') as c: thir=json.load(c)
            with open('StreamingHistory3.json') as d: four=json.load(d)
            with open('StreamingHistory4.json') as e: five=json.load(e)
            with open('StreamingHistory5.json') as f: six=json.load(f)
            with open('StreamingHistory6.json') as g: seven=json.load(g)
            with open('StreamingHistory7.json') as h: eight=json.load(h)
            with open('StreamingHistory8.json') as i: nine=json.load(i)
            with open('StreamingHistory9.json') as j: ten=json.load(j)
            datas=fir+sec+thir+four+five+six+seven+eight+nine+ten
    def questions(self):
        global nof,save,name,top,order,artno,wantmonth,period,mtop,morder,porder
        morder=0
        nof=int(input('Enter the number of your json files: '))
        save=input("Do you want to save your results in word file? yes or no: ")
        if save in ['yes','Yes']: name=input("Enter the file name you want to save as: ")
        top=int(input('Enter the number of songs shown on your top chart: '))
        order=int(input('Enter the way you want to sort out your chart\n(1 for total streaming time of the song\n2 for number of times of streaming the song): '))
        artno=int(input('Enter the number of artists you want on your top artists: '))
        wantmonth=str(input('Do you want to have your monthly charts? yes or no: '))
        if wantmonth in ['yes', 'Yes']:
            period=input('Enter specific period in the format YYYY/MM-YYYY/MM: ')
            mtop=int(input('Enter the number of songs shown on your monthly top chart: '))
            morder=int(input('Enter the way you want to sort out your chart\n(1 for total streaming time of the song\n2 for number of times of streaming the song): '))
        porder=int(input('Enter the way you want to sort out your playlist\n(1 for total streaming time of the playlist\n2 for total streaming time over total song duration of the playlist): '))
    def UltimateSorting(self):
        global ulti
        ulti={}
        for data in datas:
            artist,song,dur=data['artistName'],data['trackName'],data['msPlayed']
            if dur!=0:
                if artist in ulti.keys():
                    if song in ulti[artist].keys(): ulti[artist][song]=ulti[artist][song]+dur      #old artist old song
                    else: ulti[artist][song]=dur                                                   #old artist new song
                else: ulti[artist]={song:dur}                                                      #new artist
    def GetInfo(self):
        ultii=deepcopy(ulti)
        for artist in ultii:
            for song in ultii[artist]: spotify.search(song,artist)
        [ulti.pop(i) for i in [artist for artist in ulti if len(ulti[artist])==0]]
    def rettrash(self,n): return [tup[n] for tup in trash]        #0:artist, 1:song
mis=Miscellaneous()
class Month:
    def retmonth(self,a):
        self.a=a
        monthdict={'01':'January','02':'February','03':'March','04':'April','05':'May','06':'June','07':'July','08':'August','09':'September','10':'October','11':'November','12':'December'}
        for i in monthdict.keys():
            if a==i: return monthdict.get(i)
    def SortMonth(self,yr,month):
        self.yr,self.month=yr,month
        global multi
        multi={}
        for data in datas:
            if data['endTime'][2:4] in self.yr and data['endTime'][5:7] in self.month:
                artist,song,dur=data['artistName'],data['trackName'],data['msPlayed']
                if dur!=0:
                    if artist in multi.keys():
                        if song in multi[artist].keys(): multi[artist][song]=(multi[artist][song][0]+dur,)
                        else: multi[artist][song]=(dur,)
                    else: multi[artist]={song:(dur,)}
        [multi[tup[0]].pop(tup[1]) for tup in trash if multi.get(tup[0])!=None if multi[tup[0]].get(tup[1])!=None]
        [multi.pop(i) for i in [artist for artist in multi if len(multi[artist])==0]]
    def monthloop(self,start,end,yrshort,yrlong):
        self.start,self.end,self.yrshort,self.yrlong=start,end,yrshort,yrlong
        monthlist=['01','02','03','04','05','06','07','08','09','10','11','12']
        for i in monthlist[self.start:self.end]:
            self.SortMonth([self.yrshort],[i])
            tc.sorttop(multi)
            if len(multi)!=0: tc.presentation(mtop,morder,f' in {self.retmonth(i)} {self.yrlong}')
            else: print(f'\nYour top chart in {self.retmonth(i)} {self.yrlong} is not available.')
            multi.clear()
    def MonthlyChart(self):
        monthlist=['01','02','03','04','05','06','07','08','09','10','11','12']
        if int(period[10:12])-int(period[2:4])==0: self.monthloop(monthlist.index(period[5:7]),monthlist.index(period[13:15])+1,period[2:4],period[:4])
        for i in range(1,10):
            if int(period[10:12])-int(period[2:4])==i:
                self.monthloop(monthlist.index(period[5:7]),len(monthlist),period[2:4],period[:4])
                for j in range(1,i): self.monthloop(0,len(monthlist),str(int(period[2:4])+j),int(period[:4])+j)
                self.monthloop(0,monthlist.index(period[13:15])+1,period[10:12],period[8:12])
                break
mon=Month()
class Season:
    def rettup(self,tup,msg): return list([msg+item if len(item)!=1 else msg+'0'+item for item in tup])
    def ret(self,array,a,b): return [month[a:b] for month in array]
    def quickcheck(self,arr): return [True if i[-2:] == '12' else False for i in arr][0]
    def givszns(self,array):
        self.array=array
        for month in self.array:
            if month[-2:] in ['03','04','05']: return '20'+month[:2]+' Spring'
            elif month[-2:] in ['06','07','08']: return '20'+month[:2]+' Summer'
            elif month[-2:] in ['09','10','11']: return '20'+month[:2]+' Autumn'
            elif month[-2:] in ['12','01','02']:
                if not self.quickcheck(self.array): return '20'+str(int(month[:2])-1)+' Winter'
                else: return '20'+month[:2]+' Winter'
    def rettotaltime(self,dic):
        self.dic=dic
        coun=0
        for dict in self.dic: coun+=dict['total time']
        return coun
    def GetPeriod(self):
        global allperiod,year
        year,month1,month2=[],[],[]
        for data in datas:
            if data['trackName'] not in mis.rettrash(1):
                if len(year)==0: year.append(data['endTime'][2:4])
                elif data['endTime'][2:4]!=year[-1]: year.append(data['endTime'][2:4])
        if len(year)==1:
            for data in datas:
                if len(month1)==0: month1.append(data['endTime'][5:7])
                elif len(month1)==1:
                    if int(data['endTime'][5:7])>int(month1[-1]): month1.append(data['endTime'][5:7])
                    elif int(data['endTime'][5:7])<int(month1[-1]):
                        month1.pop(-1)
                        month1.append(data['endTime'][5:7])
                elif len(month1)>=2:
                    if int(data['endTime'][5:7])>int(month1[-1]):
                        month1.pop(-1)
                        month1.append(data['endTime'][5:7])
            allperiod=year[0]+'/'+month1[0]+'-'+year[0]+'/'+month1[1]
        elif len(year)>=2:
            for data in datas:
                if data['endTime'][2:4]==year[0]:
                    if len(month1)==0: month1.append(data['endTime'][5:7])
                    elif int(data['endTime'][5:7])<int(month1[-1]):
                        month1.pop(-1)
                        month1.append(data['endTime'][5:7])
                elif data['endTime'][2:4]==year[-1]:
                    if len(month2)==0: month2.append(data['endTime'][5:7])
                    elif int(data['endTime'][5:7])>int(month2[-1]):
                        month2.pop(-1)
                        month2.append(data['endTime'][5:7])
            allperiod=year[0]+'/'+month1[0]+'-'+year[-1]+'/'+month2[0]
    def SortUltiPeri(self):
        global ultiperiods
        #Get in dict
        periods={}
        startmonth=allperiod[4] if allperiod[3]=='0' else allperiod[3:5]
        endmonth=allperiod[10] if allperiod[9]=='0' else allperiod[9:11]
        if len(year)==1: periods[year[0]]=[str(month) for month in range(int(startmonth),int(endmonth)+1)]
        elif len(year)>=2:
            periods[year[0]]=[str(month) for month in range(int(startmonth),13)]
            for i in range(1,len(year)-1): periods[year[i]]=[str(month) for month in range(1,13)]
            periods[year[-1]]=[str(month) for month in range(1,int(endmonth)+1)]
        #Separate the months with /
        realperiods=deepcopy(periods)
        for yr in periods:
            count=0
            for ind, month in enumerate(periods[yr]):
                if month in ['3','6','9','12']:
                    realperiods[yr].insert(ind+count,'/')
                    count+=1
        #Add those months to a list
        ultiperiods=[]
        for yrind, yr in enumerate(realperiods):
            a=[]
            for ind, item in enumerate(realperiods[yr]):
                if item!='/':
                    a.append(item)
                    if item==realperiods[yr][-1]:
                        if len(a)!=0: ultiperiods.append(self.rettup(a, yr))
                else:
                    if len(a)!=0:
                        ultiperiods.append(self.rettup(a, yr))
                        a.clear()
        #Combine 12,1,2
        uultiperiods=deepcopy(ultiperiods)
        ccount=0
        for i in range(len(uultiperiods)):
            for month in uultiperiods[i]:
                if month[-2:]=='12':
                    if i!=len(uultiperiods)-1:
                        ultiperiods.remove(uultiperiods[i])
                        ultiperiods.remove(uultiperiods[i+1])
                        ultiperiods.insert(i-ccount,uultiperiods[i]+uultiperiods[i+1])
                        ccount+=1
    def SortDict(self):
        global ultiseasondict,szntotal
        ultiseasondict,szntotal={},{}
        for arr in ultiperiods:
            mon.SortMonth(self.ret(arr,0,2),self.ret(arr,2,4))
            if len(multi)!=0:
                tc.sorttop(multi)
                if len(arr)==3: szntotal[self.givszns(arr)]=self.rettotaltime(deepcopy(sstt))
                if order==2 or morder==2: ultiseasondict[self.givszns(arr)]=deepcopy(snot)
                elif order==1 or morder==1: ultiseasondict[self.givszns(arr)]=deepcopy(sstt)
                if order==2 or morder==2: snot.clear()
                sstt.clear()
                multi.clear()
    def SortSZN(self):
        self.GetPeriod()
        self.SortUltiPeri()
        self.SortDict()
    def presentation(self):
        if save not in ['yes','Yes']:
            print(f'\nDuring 20{allperiod[:2]} {mon.retmonth(allperiod[3:5])} to 20{allperiod[6:8]} {mon.retmonth(allperiod[9:11])}, your sound changed with the seasons.')
            for szn in ultiseasondict:
                print(f'\n{szn}')
                [print(dict['song']) for dict in ultiseasondict[szn][:4]]
            print('\nYour total streaming time each season is:')
            [print(f'{sznn}: {round(szntotal[sznn]/(1000*60*60),1)} hours') for sznn in szntotal]
        else:
            doc.add_paragraph(f'\nDuring 20{allperiod[:2]} {mon.retmonth(allperiod[3:5])} to 20{allperiod[6:8]} {mon.retmonth(allperiod[9:11])}, your sound changed with the seasons.')
            for szn in ultiseasondict:
                doc.add_paragraph(f'\n{szn}')
                [doc.add_paragraph(dict['song']) for dict in ultiseasondict[szn][:4]]
            doc.add_paragraph('\nYour total streaming time each season is:')
            [doc.add_paragraph(f'{sznn}: {round(szntotal[sznn]/(1000*60*60),1)} hours') for sznn in szntotal]
s=Season()
trash=[]
class SpotifyAPI:
    def __init__(self,id,secret): self.id,self.secret=id,secret
    def makereq(self,msg,typee='track'):
        global r,res
        self.msg,self.typee=msg,typee
        clientcredsb64=base64.b64encode(f"{self.id}:{self.secret}".encode())
        tokenurl='https://accounts.spotify.com/api/token'
        tokendata={'grant_type': 'client_credentials'}
        tokenheader={'Authorization': f'Basic {clientcredsb64.decode()}'}
        r1=requests.post(tokenurl, data=tokendata, headers=tokenheader)
        header={'Authorization': f"Bearer {r1.json()['access_token']}"}
        endpoint="https://api.spotify.com/v1/search"
        data=urlencode({'q':self.msg, 'type': self.typee})
        url=f"{endpoint}?{data}"
        r=requests.get(url, headers=header)
        res=r.json()[f'{self.typee}s']['items']
    def search(self,song,artist):
        global searchresults,info,artistid
        self.song,self.artist=song,artist
        self.searchkw=self.song+' '+self.artist
        def trimsong(song,trim): return song[:song.find(trim)]
        self.makereq(self.searchkw)
        if len(res)!=0: ulti[self.artist][self.song]=(ulti[self.artist][self.song],res[0]['duration_ms'],res[0]['album']['name'])
        else:
            self.makereq(trimsong(self.song,'(')+' '+self.artist)
            if len(res)!=0: ulti[self.artist][self.song]=(ulti[self.artist][self.song],res[0]['duration_ms'],res[0]['album']['name'])
            else:
                self.makereq(trimsong(self.song,'-')+' '+self.artist)
                if len(res)!=0: ulti[self.artist][self.song]=(ulti[self.artist][self.song],res[0]['duration_ms'],res[0]['album']['name'])
                else:
                    ulti[self.artist].pop(self.song)
                    trash.append((self.artist,self.song))
spotify=SpotifyAPI(clientid,clientsecret)
class TopChart:
    def sorttop(self,songdict):
        global sstt,snot
        self.songdict=songdict
        if order==2 or morder==2: snot=sorted([{'artist':artist,'song':song, 'no. of times':self.songdict[artist][song][0]/ulti[artist][song][1]} for artist in self.songdict for song in self.songdict[artist]],key=lambda x: x['no. of times'],reverse=True)
        sstt=sorted([{'artist':artist,'song':song, 'total time':self.songdict[artist][song][0]} for artist in self.songdict for song in self.songdict[artist]],key=lambda x: x['total time'],reverse=True)
    def presentation(self,top,order,monthmsg):
        self.top,self.order,self.monthmsg=top,order,monthmsg
        def msg(acc): return f"\nYour top {self.top} (at most) most-streamed songs{self.monthmsg} according to {acc} are:"
        if save not in ['yes','Yes']:
            if self.order==1:
                print(msg('total streaming time'))
                [print(f"{ind+1}. {dict['song']} by {dict['artist']} with {round(dict['total time']/ulti[dict['artist']][dict['song']][1])} times, or {round(dict['total time']/(1000*60*60),1)} hours") for ind,dict in enumerate(sstt[:self.top])]
            elif self.order==2:
                print(msg('number of times streamed'))
                [print(f"{ind+1}. {dict['song']} by {dict['artist']} with {round(dict['no. of times'])} times, or {round(dict['no. of times']*ulti[dict['artist']][dict['song']][1]/(1000*60*60),1)} hours") for ind,dict in enumerate(snot[:self.top])]
        else:
            if self.order==1:
                doc.add_paragraph(msg('total streaming time'))
                [doc.add_paragraph(f"{ind+1}. {dict['song']} by {dict['artist']} with {round(dict['total time']/ulti[dict['artist']][dict['song']][1])} times, or {round(dict['total time']/(1000*60*60),1)} hours") for ind,dict in enumerate(sstt[:self.top])]
            elif self.order==2:
                doc.add_paragraph(msg('number of times streamed'))
                [doc.add_paragraph(f"{ind+1}. {dict['song']} by {dict['artist']} with {round(dict['no. of times'])} times, or {round(dict['no. of times']*ulti[dict['artist']][dict['song']][1]/(1000*60*60),1)} hours") for ind,dict in enumerate(snot[:self.top])]
tc=TopChart()
class TopArtist:
    def sorttop(self):
        global artisttotaltime,satt,tsftat,tsftanot
        artisttotaltime=[]
        for artist in ulti:
            count=0
            for song in ulti[artist]: count+=ulti[artist][song][0]
            artisttotaltime.append({'artist':artist,'total time':count})
        satt=sorted(artisttotaltime,key=lambda x: x['total time'],reverse=True)
        #Top song from each top artist
        if order==2:
            tsftanot={}
            for dict in satt[:artno]: tsftanot[dict['artist']]=[songdict['song'] for songdict in sorted([{'song':song,'no. of times':ulti[dict['artist']][song][0]/ulti[dict['artist']][song][1]} for song in ulti[dict['artist']]],key=lambda x:x['no. of times'],reverse=True)]
        if order==1:
            tsftat={}
            for dict in satt[:artno]: tsftat[dict['artist']]=[songdict['song'] for songdict in sorted([{'song':song,'total time':ulti[dict['artist']][song][0]} for song in ulti[dict['artist']]],key=lambda x: x['total time'],reverse=True)]
    def presentation(self):
        def display(list,printt):
            if printt:
                for i,artist in enumerate(list):
                    print(f"\n{i+1}. {artist}:")
                    if len(list[artist])<=10: [print(f"{ind+1}. {song} with {round(ulti[artist][song][0]/ulti[artist][song][1])} times, or {round(ulti[artist][song][0]/(1000*60*60),1)} hours") for ind,song in enumerate(list[artist])]
                    else: [print(f"{ind+1}. {song} with {round(ulti[artist][song][0]/ulti[artist][song][1])} times, or {round(ulti[artist][song][0]/(1000*60*60),1)} hours") for ind,song in enumerate(list[artist][:10])]
            else:
                for i,artist in enumerate(list):
                    doc.add_paragraph(f"\n{i+1}. {artist}:")
                    if len(list[artist])<=10: [doc.add_paragraph(f"{ind+1}. {song} with {round(ulti[artist][song][0]/ulti[artist][song][1])} times, or {round(ulti[artist][song][0]/(1000*60*60),1)} hours") for ind,song in enumerate(list[artist])]
                    else: [doc.add_paragraph(f"{ind+1}. {song} with {round(ulti[artist][song][0]/ulti[artist][song][1])} times, or {round(ulti[artist][song][0]/(1000*60*60),1)} hours") for ind,song in enumerate(list[artist][:10])]
        if save not in ['yes','Yes']:
            print(f"\nYour top {artno} (at most) artists are:")
            [print(f"{ind+1}. {dict['artist']} with {round(dict['total time']/(1000*60*60))} hours") for ind,dict in enumerate(satt[:artno])]
            print(f"\nYour top 10 (at most) songs from your top {artno} artists are: ")
            if order==1: display(tsftat,True)
            else: display(tsftanot,True)
        else:
            doc.add_paragraph(f"\nYour top {artno} (at most) artists are:")
            [doc.add_paragraph(f"{ind+1}. {dict['artist']} with {round(dict['total time']/(1000*60*60))} hours") for ind,dict in enumerate(satt[:artno])]
            doc.add_paragraph(f"\nYour top 10 (at most) songs from your top {artno} artists are: ")
            if order==1: display(tsftat,False)
            else: display(tsftanot,False)
        if order==2: snot.clear()
        sstt.clear()
ta=TopArtist()
class TopAlbum:
    def SortnPresent(self):
        global albums
        albums={}
        for artist in ulti:
            for song in ulti[artist]:
                tup=ulti[artist][song]
                if song not in mis.rettrash(1):
                    if tup[2] in albums.keys(): albums[tup[2]]=(albums[tup[2]][0]+tup[0],artist)
                    else: albums[tup[2]]=(tup[0],artist)
        sortedalbum=sorted(albums,key=lambda x:albums[x][0],reverse=True)
        if save not in ['yes','Yes']:
            print('\nYour top 5 (at most) albums are: ')
            [print(f'{ind+1}. {album} by {albums[album][1]} for {round(albums[album][0]/(1000*60*60),1)} hours') for ind,album in enumerate(sortedalbum[:5])]
        else:
            doc.add_paragraph('Your top 5 (at most) albums are: ')
            [doc.add_paragraph(f'{ind+1}. {album} by {albums[album][1]} for {round(albums[album][0]/(1000*60*60),1)} hours') for ind,album in enumerate(sortedalbum[:5])]
tal=TopAlbum()
class TopPlaylist:
    def SortPlaylist(self):
        global plnot,plt,playlists
        playlists={}
        with open('Playlist1.json') as a: pdatas=json.load(a)
        for pldict in pdatas['playlists']:
            plname=pldict['name']
            for trackdict in pldict['items']:
                songg=trackdict['track']['trackName']
                singer=trackdict['track']['artistName']
                if singer in ulti.keys():                      #Artist appears on the data
                    if songg in ulti[singer].keys():           #Song of the artist appears on the data
                        if plname in playlists.keys(): playlists[plname]={'total time':playlists[plname]['total time']+ulti[singer][songg][0],'cumulative songdur':playlists[plname]['cumulative songdur']+ulti[singer][songg][1]}
                        else: playlists[plname]={'total time':ulti[singer][songg][0],'cumulative songdur':ulti[singer][songg][1]}
                else:                                          #Artist cooperation that actually appears on data, Didn't even appear on data
                    for artist in ulti:
                        if songg in ulti[artist].keys():       #The song actually appears on data
                            if plname in playlists.keys(): playlists[plname]={'total time':playlists[plname]['total time']+ulti[artist][songg][0],'cumulative songdur':playlists[plname]['cumulative songdur']+ulti[artist][songg][1]}
                            else: playlists[plname]={'total time':ulti[artist][songg][0],'cumulative songdur':ulti[artist][songg][1]}
        if porder==2: plnot=sorted(playlists,key=lambda x: playlists[x]['total time']/playlists[x]['cumulative songdur'],reverse=True)
        if porder==1: plt=sorted(playlists,key=lambda x: playlists[x]['total time'],reverse=True)
    def presentation(self):
        def msg(acc): return f'\nYour top 5 (at most) playlists according to {acc} are:'
        if save not in ['yes', 'Yes']:
            if porder==2:
                print(msg('equivalent number of times streamed'))
                [print(str(ind+1)+'. '+pl+' for '+str(round(playlists[pl]['total time']/playlists[pl]['cumulative songdur']))+' times') for ind,pl in enumerate(plnot[:5])]
            elif porder==1:
                print(msg('total streaming time'))
                [print(str(ind+1)+'. '+pl+' for '+str(round(playlists[pl]['total time']/(1000*60*60),1))+' hours') for ind,pl in enumerate(plt[:5])]
        else:
            if porder==2:
                doc.add_paragraph(msg('equivalent number of times streamed'))
                [doc.add_paragraph(str(ind+1)+'. '+pl+' for '+str(round(playlists[pl]['total time']/playlists[pl]['cumulative songdur']))+' times') for ind,pl in enumerate(plnot[:5])]
            elif porder==1:
                doc.add_paragraph(msg('total streaming time'))
                [doc.add_paragraph(str(ind+1)+'. '+pl+' for '+str(round(playlists[pl]['total time']/(1000*60*60),1))+' hours') for ind,pl in enumerate(plt[:5])]
tp=TopPlaylist()
class Songofthehour:
    def SortHour(self):
        global topsongofhour
        hours,topsongofhour={},{}
        for data in datas:
            time,artist,song,dur=data['endTime'],data['artistName'],data['trackName'],data['msPlayed']
            if dur!=0:
                if song not in mis.rettrash(1):
                    if time[11:13] in hours.keys():
                        if song in hours[time[11:13]].keys():
                            if artist==hours[time[11:13]].get(song)[0]: hours[time[11:13]][song]=[artist,hours[time[11:13]][song][1]+dur]
                            else:
                                if song+' ' not in hours[time[11:13]].keys(): hours[time[11:13]][song+' ']=[artist,dur]
                                else: hours[time[11:13]][song+' ']=[artist,hours[time[11:13]][song+' '][1]+dur]
                        else: hours[time[11:13]][song]=[artist,dur]
                    else: hours[time[11:13]]={song:[artist,dur]}
        shours=dict(sorted(hours.items()))
        for hour in shours:
            kk={song:[shours[hour][song][0],shours[hour][song][1]] for song in shours[hour]}
            key=sorted(kk,key=lambda x: kk[x][1],reverse=True)[0]
            topsongofhour[hour]={key:[hours[hour][key][0],hours[hour][key][1]]}
    def presentation(self):
        if save not in ['yes','Yes']:
            print('\nThis is your suggested song of the hour (in UTC). Enjoy')
            [print(f'{hour}:00 - {list(topsongofhour[hour].keys())[0]} by {topsongofhour[hour][list(topsongofhour[hour].keys())[0]][0]}') for hour in topsongofhour]
        else:
            doc.add_paragraph('\nThis is your suggested song of the hour (in UTC). Enjoy')
            [doc.add_paragraph(f'{hour}:00 - {list(topsongofhour[hour].keys())[0]} by {topsongofhour[hour][list(topsongofhour[hour].keys())[0]][0]}') for hour in topsongofhour]
soth=Songofthehour()

def main():
    mis.questions()
    mis.readjson()
    mis.UltimateSorting()
    mis.GetInfo()
    tc.sorttop(ulti)
    tc.presentation(top,order,'')
    ta.sorttop()
    ta.presentation()
    if wantmonth in ['yes','Yes']: mon.MonthlyChart()
    tal.SortnPresent()
    tp.SortPlaylist()
    tp.presentation()
    s.SortSZN()
    s.presentation()
    soth.SortHour()
    soth.presentation()


if __name__=='__main__':
    try: main()
    except FileNotFoundError: print('Please enter the correct number of json files!')
    except (ConnectionAbortedError,ConnectionError,ConnectionRefusedError,ConnectionResetError,UnicodeDecodeError,UnicodeEncodeError,UnicodeError,TimeoutError): print('Something is wrong with the connection. Please try it again.')
    except Exception: print('Something went wrong. Please try it again.')

global wantmonth,save,name
if save in ['yes','Yes']: doc.save(name+'.docx')