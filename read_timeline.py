import win32com.client
import tweepy
import configparser
import time
import sys


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
        read_timeline(latestid)
    else:
        print("There were no new tweets during 5.5s.")
        time.sleep(5.5)
        read_timeline(latestid)


def reAuth(api):
    global results
    while(True):
        print("TwitterAPI could not authenticate your account.\nGet your ConsumerKey,ConsumerSecret,AccessToken,AccessTokenSecret!")
        print("Detail:https://developer.twitter.com/en/apply-for-access\n")
        CK = input("input your ConsumerKey:")
        CS = input("input your ConsumerSecret:")
        AT = input("input your AccessToken:")
        AS = input("input your AccessTokenSecret:")
        config = configparser.ConfigParser()
        section = "TwitterAPI"
        config.add_section(section)
        config.set(section, "CK", CK)
        config.set(section, "CS", CS)
        config.set(section, "AT", AT)
        config.set(section, "AS", AS)
        auth = tweepy.OAuthHandler(CK, CS)
        auth.set_access_token(AT, AS)
        api = tweepy.API(auth)
        try:
            results = api.home_timeline(count=2)
            print("Authentication　Successful!")
            with open("config.ini", "w")as f:
                config.write(f)
            print("Wrote your CK,CS,AT,AS to config.ini")
            select = input("Restart to erase prompt?(y/n):")
            if select == "y":
                sys.exit()
            else:
                print("Setting finished!")
                break
        except Exception as e:
            print(e)
            print(CK)
            print(CS)
            print(AT)
            print(AS)
            print("Authentication failed…Please retry.")
    return api


if __name__ == '__main__':
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
    try:
        read_timeline(None)
    except Exception as e:
        print(e)
        api = reAuth(api)
        read_timeline(None)
