import win32com.client
import tweepy
import configparser
import time

config = configparser.ConfigParser()
config.read("config.ini")
section = "TwitterAPI"
try:
    CK = config.get(section, "ck")
    CS = config.get(section, "cs")
    AT = config.get(section, "at")
    AS = config.get(section, "as")
except:
    CK, CS, AT, AS = "", "", "", ""

auth = tweepy.OAuthHandler(CK, CS)
auth.set_access_token(AT, AS)

api = tweepy.API(auth)


sapi = win32com.client.Dispatch("SAPI.SpVoice")
cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
cat.SetID(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", True)


def speak(word):
    v = [t for t in cat.EnumerateTokens() if t.GetAttribute("Name")
         == "Microsoft Sayaka"]
    if v:
        oldv = sapi.Voice
        sapi.Voice = v[0]
        sapi.Speak(word)
        sapi.Voice = oldv


def read_timeline(latestid):
    start = time.time()
    results = api.home_timeline(count=2, since_id=latestid)
    if (len(results) > 0):
        latestid = results[0].id
        for status in results:
            print(status.user.name+':'+status.text)
            speak(status.text)
        elapsed = time.time()-start
        if(elapsed < 5.5):
            time.sleep(5.5-elapsed)
        else:
            read_timeline(latestid)
    else:
        time.sleep(5.5)
        read_timeline(latestid)


if __name__ == '__main__':
    read_timeline(None)
