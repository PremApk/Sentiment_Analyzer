from tkinter import filedialog
from tkinter import *
from PIL import ImageTk,Image
import pandas as pd
import matplotlib.pyplot as plt
from nltk.sentiment.vader import SentimentIntensityAnalyzer
import xlsxwriter
from textblob import TextBlob
import tweepy
filename=""
def close_windowf ():
    obj.destroy()
    def browse():
        global filename
        filename=filedialog.askopenfilename()
        text.insert(0.0,filename)

    def close_window ():
        obj1.destroy()
        print(filename)
        f1=open(filename, "rb")
        workbook = xlsxwriter.Workbook('new.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'data')

        flag = 0
        flag1 = 0
        flag2 = 0
        char1 = ""
        i = 0
        j = 1
        neutral = 0
        positive = 0
        negative = 0
        sentence = []
        c = str(f1.read(1))
        input_1=eval(input("Enter the number(0 for Heading ):"))
        while c != "b''":
            c = str(f1.read(1))
            char = c[2:3]
            if c == -1:
                break
            if (char == "<")and(flag == 0):
                flag = 1
            elif (char == "d")and(flag == 1):
                flag = 2
            elif(char == "i")and(flag == 2):
                flag = 3
            elif(char == "v")and(flag == 3):
                flag = 4
            elif(char == ">" and flag == 4 and input_1==1)or(char ==" " and flag==4 and input_1==0):
                if(char==">" and input_1==1):
                    flag = 5
                else:
                    flag=11
            elif ((char=="c" and flag==11 and input_1==0))or((char != "<" and flag == 5 and input_1==1)and((char.isalpha() == True)or(char == " ")or(char.isdigit() == True)or(char == ":")or(char == "."))):
                if(flag==11 and input_1==0):
                    flag=6
                else:
                    char1 = char1 + char
                    flag1 = 1
            elif(char=="l")and(flag==6)and input_1==0:
                flag=7
            elif (char == "a") and (flag == 7):
                flag = 15
            elif (char == "s") and (flag == 15):
                flag = 16
            elif (char == "s") and (flag == 16):
                flag = 17
            elif (char == "=") and (flag == 17):
                flag = 18
            elif (char == "\"") and (flag == 18):
                flag = 19
            elif(char == "_")and(flag==19):
                flag=8
            elif(char == "3")and(flag==8):
                flag=13
            elif((char == "-")and(flag==13))or(flag==9)or(flag==12):
                if(char == "-")and(flag==8):
                    flag=9
                elif(char==">" and flag==9):
                    flag=12
                elif(flag==12):
                    if((char.isalpha() == True)or(char == " ")or(char.isdigit() == True)or(char == ":")or(char == "."))and(char!="<"):
                        char1 = char1 + char
                        flag1 = 1
                    elif(char=="<"):
                        flag=14
                else:
                    flag=9
            else:
                flag = 0
                if flag1 == 1:
                    flag1 = 0
                    flag2 = 1
            if flag2 == 1:
                worksheet.write(j,i,char1)
                j = j+1
                char1 = ""
                flag2 = 0
        workbook.close()
        f1.close()

        file = 'new.xlsx'
        data = pd.ExcelFile(file)
        parse_data = data.parse(data.sheet_names[0])
        filter_data = list(parse_data['data'])
        sia = SentimentIntensityAnalyzer()

        for i in filter_data:
            sentence = [i.find("Jan"),
                i.find("Feb"),
                i.find("Mar"),
                i.find("Apr"),
                i.find("May"),
                i.find("Jun"),
                i.find("Jul"),
                i.find("Aug"),
                i.find("Sep"),
                i.find("Oct"),
                i.find("Nov"),
                i.find("Dec")
                ]
            if len(set(sentence)) == 1:
                final = sia.polarity_scores(i)
                print(i)
                print(final)
                if final['compound'] >= 0.05:
                    positive += 1
                elif final['compound'] <= -0.05:
                    negative += 1
                else:
                    neutral += 1
        print("Positive", positive)
        print("Negative", negative)
        print("Neutral", neutral)

        labels = ["Positive", "Negative", "Neutral"]
        sizes= [positive, negative, neutral]
        colors = ["yellowgreen", "lightcoral", "gold"]
        explode = (0.1, 0, 0)
        plt.pie(sizes, explode=explode, labels=labels, colors=colors, autopct="%1.1f%%", shadow=True)
        plt.tight_layout()
        plt.axis("equal")
        plt.show()

    obj1=Tk()
    obj1.title("Sentamental Analysis")
    photo=PhotoImage(file="download.png")
    obj1.iconphoto(False,photo)
    obj1.resizable(0, 0) 

    canvas=Canvas(obj1,width=400,height=300)
    canvas.pack()
    image=ImageTk.PhotoImage(Image.open("download5.jpg"))
    canvas.create_image(0,0,anchor=NW,image=image)
    image1=ImageTk.PhotoImage(Image.open("download2.jpg"))
    canvas.create_image(0,138,anchor=NW,image=image1)
    text=Text(obj1,width=30,height=1,wrap=WORD)
    button2=canvas.create_window(65, 100, anchor='nw', window=text)
    myButton = Button(obj1,text="OPEN",borderwidth=3,command=browse,bg="medium blue",fg="ghostwhite")
    button=canvas.create_window(320, 96, anchor='nw', window=myButton)
    myButton1 = Button(obj1,text="ANALYSE",borderwidth=3,command=close_window,bg="medium blue",fg="ghostwhite")
    button1=canvas.create_window(170, 150, anchor='nw', window=myButton1)

    width=400
    height=300
    screenw=obj1.winfo_screenwidth()
    screenh=obj1.winfo_screenheight()
    x=(screenw/2)-(width/2)
    y=(screenh/2)-(height/2)
    obj1.geometry("%dx%d+%d+%d" % (width,height,x,y))

    obj1.mainloop()



def close_windowt ():
    obj.destroy()
    def close_window ():
        keyword=text.get(0.0,END)
        count=text1.get(0.0,END)
        obj1.destroy()
        count=int(count)
        consumer_key = 'XXXXXXXXXXXXXXXXXXXXX'
        consumer_secret_key = 'XXXXXXXXXXXXXXXXXXXXXXXXXXX'
        access_key = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXX'
        access_secret_key = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

        def percentage(part, whole):
            return 100 * float(part)/float(whole)

        positive = 0
        negative = 0
        neutral = 0
        polarity = 0

        auth = tweepy.OAuthHandler(consumer_key, consumer_secret_key)
        auth.set_access_token(access_key, access_secret_key)
        api = tweepy.API(auth)
        #keyword = input("Enter the Hashtag/Keyword :")
        #count = input("Enter the searchcount")
        tweets = api.search(keyword, count=count)
        for i in tweets:
            print(i.text)
            analysis = TextBlob(i.text)
            polarity += analysis.sentiment.polarity
            print(polarity)
            if(analysis.sentiment.polarity == 0):
                neutral += 1
            elif(analysis.sentiment.polarity < 0):
                negative += 1
            else:
                positive += 1

        positive = percentage(positive, 100)
        negative = percentage(negative, 100)
        neutral = percentage(neutral, 100)

        positive = format(positive, '.2f')
        negative = format(negative, '.2f')
        neutral = format(neutral, '.2f')

        labels = ['Positive ['+str(positive)+'%]', 'Neutral ['+str(neutral)+'%]', 'Negative ['+str(negative)+'%]']
        sizes = [positive,neutral,negative]
        colors = ['yellowgreen','gold','red']
        patches,texts = plt.pie(sizes, colors=colors, startangle=90)
        plt.legend(patches, labels, loc="best")
        plt.axis('equal')
        plt.tight_layout()
        plt.show()

    obj1=Tk()
    obj1.title("Sentamental Analysis")
    photo=PhotoImage(file="download.png")
    obj1.iconphoto(False,photo)
    obj1.resizable(0, 0) 

    canvas=Canvas(obj1,width=400,height=300)
    canvas.pack()
    image=ImageTk.PhotoImage(Image.open("download.jpg"))
    canvas.create_image(0,0,anchor=NW,image=image)
    image1=ImageTk.PhotoImage(Image.open("download2.jpg"))
    canvas.create_image(0,138,anchor=NW,image=image1)
    text=Text(obj1,width=30,height=1,wrap=WORD)
    button2=canvas.create_window(110, 60, anchor='nw', window=text)
    text1=Text(obj1,width=5,height=1,wrap=WORD)
    button=canvas.create_window(110, 100, anchor='nw', window=text1)
    myButton3 = Button(obj1,text="KEYWORD",borderwidth=3,command=close_window,bg="deep sky blue",fg="black")
    button3=canvas.create_window(15, 57, anchor='nw', window=myButton3)
    myButton4 = Button(obj1,text="COUNT",borderwidth=3,command=close_window,bg="deep sky blue",fg="black")
    button4=canvas.create_window(15, 97, anchor='nw', window=myButton4)
    myButton1 = Button(obj1,text="ANALYSE",borderwidth=3,command=close_window,bg="deep sky blue",fg="black")
    button1=canvas.create_window(170, 150, anchor='nw', window=myButton1)

    width=400
    height=300
    screenw=obj1.winfo_screenwidth()
    screenh=obj1.winfo_screenheight()
    x=(screenw/2)-(width/2)
    y=(screenh/2)-(height/2)
    obj1.geometry("%dx%d+%d+%d" % (width,height,x,y)) 

    obj1.mainloop()


obj=Tk()
obj.title("Sentamental Analysis")
photo=PhotoImage(file="download.png")
obj.iconphoto(False,photo)
obj.resizable(0, 0) 


canvas=Canvas(obj,width=400,height=300) 
canvas.pack()
image=ImageTk.PhotoImage(Image.open("download3.jpg"))
canvas.create_image(0,0,anchor=NW,image=image)

myButton3 = Button(obj,text="FACEBOOK",borderwidth=2,command=close_windowf,bg="medium blue",fg="ghostwhite")
button3=canvas.create_window(210, 130, anchor='nw', window=myButton3)
myButton4 = Button(obj,text="TWITTER",borderwidth=2,command=close_windowt,bg="deep sky blue",fg="black")
button4=canvas.create_window(210, 100, anchor='nw', window=myButton4)



width=image.width()
height=image.height()
screenw=obj.winfo_screenwidth()
screenh=obj.winfo_screenheight()
x=(screenw/2)-(width/2)
y=(screenh/2)-(height/2)
obj.geometry("%dx%d+%d+%d" % (width,height,x,y))

 

obj.mainloop()
