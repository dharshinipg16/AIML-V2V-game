# -*- coding: utf-8 -*-


import speech_recognition as sr
import pyttsx3

# to read xls file
import xlrd 
loc = ("quizmps.xls")
wb = xlrd.open_workbook(loc) 
score=0
qn=0
limit=5

# this is to call test to speech 
engine = pyttsx3.init()
language = 'en'
mytext="start"
r = sr.Recognizer()

rate=engine.getProperty("rate")
engine.setProperty("rate",130)
volume=engine.getProperty("volume")
engine.setProperty("volume",1.0)

index=0
# methos to say sorry
def sorry():
    print("sorry I didn't get you... speak again...")
    engine.say("Sorry I didn't get you speak again")
    engine.runAndWait()
# methos to say sorry for google connection
def sorry1():
    print("sorry I didn't get you... check your google connectivity...")
    engine.say("Sorry I didn't get you check your google connectivity")
    engine.runAndWait()
# method to call main menu
def voicemenu():
    print("\tWELCOME TO MY  QUIZ USING MACHINE LEARNING ")
    engine.say("WELCOME TO MY  QUIZ USING MACHINE LEARNING")
    print("\t\tSAY SPORTS for SPORTS")
    engine.say("SAY SPORTS for SPORTS")
    print("\t\tSAY FILM for MOVIES")
    engine.say("SAY FILM for MOVIES")
    print("\t\tSAY FOOD for FOOD AND CULTURE")
    engine.say("SAY FOOD for FOOD AND CULTURE")
    print("\t\tSAY POLITICS for INDIAN POLITICS")
    engine.say("SAY POLITICS for INDIAN POLITICS")
    print("\t\tSAY SCIENCE for SCIENCE")
    engine.say("SAY SCIENCE for SCIENCE")
    print("\t\tSAY DONE TO QUIT GAME")
    engine.say("SAY DONE TO QUIT GAME")
    engine.runAndWait()


voicemenu()
flag=True
with sr.Microphone() as source:
    while(flag):
        try:
            print("\n\tPlease give your choice...")
            engine.say("Please give your choice")
            engine.runAndWait()
            #r.adjust_for_ambient_noise(source)
            audio = r.listen(source)
            mytext=r.recognize_google(audio)
            print("you said:{}".format(mytext))
            
            ########################################
            if(mytext=="Sports"):
                engine.say("your choice is sports")
                engine.runAndWait()
                sheet = wb.sheet_by_index(0)
            
                flag=True
                with sr.Microphone() as source:
    
                    while(qn<limit):
        
                        try:
                            # t1 is a temp variable to get qn and ans
                            t1=sheet.cell_value(qn,1)
                            getqn=str(t1)
                            print("\nYour Question is: "+getqn)
                            engine.say(getqn)
                            engine.runAndWait()
                            #r.adjust_for_ambient_noise(source)
                            print("\nGive your Answer now ")
                            engine.say("Give your Answer now")
                            engine.runAndWait()
                            audio1 = r.listen(source)
                            mytext1=r.recognize_google(audio1)
                            print("you said:{}".format(mytext1))
                            mytext1=mytext1.lower()
                            t1=sheet.cell_value(qn,2)
                            ans=str(t1)
                            ans=ans.lower()
                            if(ans==mytext1):
                                score+=10
                                print("\nYou are correct... ")
                                engine.say("You are correct")
                                engine.runAndWait()
                                    
                                
                            else:
                                print("\nSorry the Answer is: "+ans)
                                t2="Sorry the Answer is "+ans
                                engine.say(t2)
                                engine.runAndWait()
                            qn+=1
                            print("Question number:",qn)
                    
                        except:
                            if(qn==limit):
                                flag=False 
                            else:    
                                sorry1()
                        #qn=qn-1
                
                
            #############################################  
            elif(mytext=="film") or (mytext=="movie"):
                engine.say("your choice is movies")
                engine.runAndWait()
                sheet = wb.sheet_by_index(1)
            
                flag=True
                with sr.Microphone() as source:
    
                    while(qn<limit):
        
                        try:
                            # t1 is a temp variable to get qn and ans
                            t1=sheet.cell_value(qn,1)
                            getqn=str(t1)
                            print("\nYour Question is: "+getqn)
                            engine.say(getqn)
                            engine.runAndWait()
                            #r.adjust_for_ambient_noise(source)
                            print("\nGive your Answer now ")
                            engine.say("Give your Answer now")
                            engine.runAndWait()
                            audio1 = r.listen(source)
                            mytext1=r.recognize_google(audio1)
                            print("you said:{}".format(mytext1))
                            mytext1=mytext1.lower()
                            t1=sheet.cell_value(qn,2)
                            ans=str(t1)
                            ans=ans.lower()
                            if(ans==mytext1):
                                score+=10
                                print("\nYou are correct... ")
                                engine.say("You are correct")
                                engine.runAndWait()
                                    
                                
                            else:
                                print("\nSorry the Answer is: "+ans)
                                t2="Sorry the Answer is "+ans
                                engine.say(t2)
                                engine.runAndWait()
                            qn+=1
            
                    
                        except:
                            if(qn==limit):
                                flag=False 
                            else:
                                sorry1()
                        #qn=qn-1
                
               
                
             ##########################################   
            elif(mytext=="food"):
                engine.say("your choice is Food and Culture")
                engine.runAndWait()   
                sheet = wb.sheet_by_index(2)
            
                flag=True
                with sr.Microphone() as source:
    
                    while(qn<limit):
        
                        try:
                            # t1 is a temp variable to get qn and ans
                            t1=sheet.cell_value(qn,1)
                            getqn=str(t1)
                            print("\nYour Question is: "+getqn)
                            engine.say(getqn)
                            engine.runAndWait()
                            #r.adjust_for_ambient_noise(source)
                            print("\nGive your Answer now ")
                            engine.say("Give your Answer now")
                            engine.runAndWait()
                            audio1 = r.listen(source)
                            mytext1=r.recognize_google(audio1)
                            print("you said:{}".format(mytext1))
                            mytext1=mytext1.lower()
                            t1=sheet.cell_value(qn,2)
                            ans=str(t1)
                            ans=ans.lower()
                            if(ans==mytext1):
                                score+=10
                                print("\nYou are correct... ")
                                engine.say("You are correct")
                                engine.runAndWait()
                                    
                                
                            else:
                                print("\nSorry the Answer is: "+ans)
                                t2="Sorry the Answer is "+ans
                                engine.say(t2)
                                engine.runAndWait()
                            qn+=1
            
                    
                        except:
                            if(qn==limit):
                                flag=False 
                            else:
                                sorry1()
                        #qn=qn-1
                
                
             ###############################################   
            elif(mytext=="politics"):
                engine.say("your choice is Indian Politics")
                engine.runAndWait()
                sheet = wb.sheet_by_index(3)
            
                flag=True
                with sr.Microphone() as source:
    
                    while(qn<limit):
        
                        try:
                            # t1 is a temp variable to get qn and ans
                            t1=sheet.cell_value(qn,1)
                            getqn=str(t1)
                            print("\nYour Question is: "+getqn)
                            engine.say(getqn)
                            engine.runAndWait()
                            #r.adjust_for_ambient_noise(source)
                            print("\nGive your Answer now ")
                            engine.say("Give your Answer now")
                            engine.runAndWait()
                            audio1 = r.listen(source)
                            mytext1=r.recognize_google(audio1)
                            print("you said:{}".format(mytext1))
                            mytext1=mytext1.lower()
                            t1=sheet.cell_value(qn,2)
                            ans=str(t1)
                            ans=ans.lower()
                            if(ans==mytext1):
                                score+=10
                                print("\nYou are correct... ")
                                engine.say("You are correct")
                                engine.runAndWait()
                                    
                                
                            else:
                                print("\nSorry the Answer is: "+ans)
                                t2="Sorry the Answer is "+ans
                                engine.say(t2)
                                engine.runAndWait()
                            qn+=1
            
                        except:
                            if(qn==limit):
                                flag=False 
                            else:
                                sorry1()
                        #qn=qn-1
                
                
                
             ##############################################   
            elif(mytext=="science"):
                engine.say("your choice is Science")
                engine.runAndWait()
                sheet = wb.sheet_by_index(4)
            
                flag=True
                with sr.Microphone() as source:
    
                    while(qn<limit):
        
                        try:
                            # t1 is a temp variable to get qn and ans
                            t1=sheet.cell_value(qn,1)
                            getqn=str(t1)
                            print("\nYour Question is: "+getqn)
                            engine.say(getqn)
                            engine.runAndWait()
                            #r.adjust_for_ambient_noise(source)
                            print("\nGive your Answer now ")
                            engine.say("Give your Answer now")
                            engine.runAndWait()
                            audio1 = r.listen(source)
                            mytext1=r.recognize_google(audio1)
                            print("you said:{}".format(mytext1))
                            mytext1=mytext1.lower()
                            t1=sheet.cell_value(qn,2)
                            ans=str(t1)
                            ans=ans.lower()
                            if(ans==mytext1):
                                score+=10
                                print("\nYou are correct... ")
                                engine.say("You are correct")
                                engine.runAndWait()
                                    
                                
                            else:
                                print("\nSorry the Answer is: "+ans)
                                t2="Sorry the Answer is "+ans
                                engine.say(t2)
                                engine.runAndWait()
                            qn+=1
            
                    
                        except:
                            if(qn==limit):
                                flag=False 
                            else:
                                 sorry1()
                        #qn=qn-1
                
                
            #################################################    
            elif(mytext=="done"):
                engine.say("Thank you bye")
                engine.runAndWait()
                flag=False
            
            ##################################################
            else:
                sorry()
        except:
            if(qn==limit):
                flag=False 
            else:
                sorry1()

print("Your Score is: "+str(score))
bye=" Your Score is "+str(score)
engine.say(bye)
engine.say("Thank you bye bye and Thanks for playing Dharshenees game")
engine.runAndWait()
