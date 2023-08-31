from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import random

# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
import json
from selenium import webdriver 
import chromedriver_autoinstaller 
import os
 
#########################################################################################
#                                                                                       #
#                SECTION FOR SAVE AND LOAD CHECK POINTS                                 #
#                                                                                       #
#########################################################################################
def saveCheckPoint(filename, dictionary):
    json_object = json.dumps(dictionary, indent=4)
    with open(filename, "w") as outfile:
        outfile.write(json_object)

def loadCheckPoint(filename):
    # Opening JSON file
    with open(filename, 'r') as openfile:        
        json_object = json.load(openfile)
    return json_object

def saveLastArticle(idnumber, link):
    filename = 'files_prepssm/lastexam.json'
    dictionary = {'last':idnumber, 'link':link}
    json_object = json.dumps(dictionary)    
    with open(filename, "w") as outfile:
        outfile.write(json_object)
#########################################################################################

#########################################################################################
#                                                                                       #
#                SECTION FOR GET DATA FROM EACH EXAM                                    #
#                                                                                       #
#########################################################################################
def getQuestionNumber(question):
    questionnumber = question.find_element(By.CLASS_NAME, 'question-title.mr-auto')
    return questionnumber.text

def getQuestionDescription(question):
                    # 'question-description.select-text ng-star-inserted'
    # classquestion = 'question-description.select-text.nnki-no-select.ng-star-inserted' Original
    classquestion = 'question-description.select-text.ng-star-inserted'
    questiondescription = question.find_element(By.CLASS_NAME, classquestion)
    dictQuestion['Question'] = [questiondescription.text]    

def getOptions(question):
    """ In each question block first get complete block with the available options,
    then iterate on each option getting only obition wiout the letter option """
    try:
        # class to get block with all the options class option1
        blockoptionsclass = 'form-nnki.form-nnki-candidate.-correction.ng-untouched.ng-pristine.ng-valid.ng-star-inserted'# block options
        blockoption = question.find_element(By.CLASS_NAME, blockoptionsclass)
    except:
        # class to get block witho all the options class option2
        blockoptionsclass = 'form-nnki.form-nnki-candidate.-correction.ng-untouched.ng-pristine.ng-star-inserted'
        blockoption = question.find_element(By.CLASS_NAME, blockoptionsclass)

    separeteopcions = 'row-form-nnki.row-answer.ng-star-inserted' # fot get options
    listoptions = blockoption.find_elements(By.CLASS_NAME, separeteopcions)
    
    for i, option in enumerate(listoptions):
        optionID = getOptionKey(option)        
        dictQuestion[optionID] = re.sub(r'([A-E]\.\s+)', '',option.text)

def getCorrectOption(question):
    """FUNCTION TO GET THE CORRECT OPTION"""
    classcorrect = 'row-form-nnki.row-answer.-answerShould.ng-star-inserted'
    correctAnwer = question.find_element(By.CLASS_NAME, classcorrect)
    
    optionID = getOptionKey(correctAnwer)    
    dictQuestion[optionID] = re.sub(r'[A-E]\.\s+', '',correctAnwer.text)#correctAnwer.text
    dictQuestion['CorrectOption'] = optionID

def getOptionKey(option):
    # FUNCTION TO EXTRACT THE TEXT IN EACH OPTION
    text = option.get_attribute('outerHTML')
    optionID = re.findall(r'([A-E])\.&nbsp;', text)[0]
    return optionID

def getCorrection(question):
    """FUNCTION TO GET EXPLENATION OF THE QUESTION"""
    blockanswer = question.find_element(By.CLASS_NAME, "card-content.-correction")
    correctanswer = blockanswer.find_element(By.CLASS_NAME, 'card-content-inside')
    dictQuestion['Correction'] = correctanswer.text            
    return correctanswer

def getQuestionLinkImage(question):
    """FUNCTION TO GET LINKS IMAGES APPERED IN QUESTION SECTION"""
    global listcolumns
    cardcontent = question.find_element(By.CLASS_NAME, "card-content")
    insidecard = cardcontent.find_element(By.CLASS_NAME,"card-content-inside")
    
    if not('QuestionImage_src' in listcolumns):
            listcolumns.append('QuestionImage_src')
    try:
        images = insidecard.find_elements(By.CLASS_NAME, "w-full.rounded.relative.cursor-pointer.ng-star-inserted")
        
        if len(images) == 0:
            dictQuestion['QuestionImage_src'] = '*'    

        for i, image in enumerate(images):
            imagelink = image.get_attribute("src")
            if i == 0:
                dictQuestion['QuestionImage_src'] = imagelink
                if not('QuestionImage_src' in listcolumns):
                    listcolumns.append('QuestionImage_src')
            else:
                dictQuestion['QuestionImage_src_{}'.format(i)] = imagelink
                if not('QuestionImage_src_{}'.format(i) in listcolumns):
                    listcolumns.append('QuestionImage_src_{}'.format(i))
    except:
        dictQuestion['QuestionImage_src'] = '*'    

def getLinkAnswer(correctanswer):
    
    #########################################
    #                                       #
    #   GET LINKS CONTAINED IN ANSWER TEXT  #
    #                                       #
    #########################################

    global listcolumns
    blockhtml = correctanswer.get_attribute('outerHTML')    
    correctionlinks = re.findall("<a\s+.+?/a>", blockhtml)

    for i, link in enumerate(correctionlinks):        
        try:
            correctionhref = re.findall(r'href="(.+?)">',link)[0]    
            keyhref = "CorrectionLink_{}_href".format(i+1)        
            dictQuestion[keyhref] = correctionhref
            if not(keyhref in listcolumns):
                listcolumns.append(keyhref)

            texto = re.findall(r'\>(.+?)\<', link)[0]
            keyText = "CorrectionLink_{}".format(i+1)    
            dictQuestion[keyText] = texto        
            if not(keyText in listcolumns):
                listcolumns.append(keyText)
        except:
            "No links"

def getLinkImageAnswer(correctanswer):
    """GET THE LINK OF IMAGES IN ANSWER SECTION"""
    global listcolumns
    imagelinkAnswer = correctanswer.find_elements(By.CLASS_NAME, "w-full.rounded.relative.cursor-pointer.ng-star-inserted")

    for i, link in enumerate(imagelinkAnswer):
        imgAnswerlink = link.get_attribute('src')
        keyImage = "CorrectionImage_{}_src".format(i+1)
        dictQuestion[keyImage] = imgAnswerlink
        
        if not(keyImage in listcolumns):
            listcolumns.append(keyImage)

def getVideosLinks():
    """GET VIDEO LINKS IN THE INITIAL PART, IT ONLY APPEAR IN SOME TEXT NO ALL"""
    try:
        videos = driver.find_elements(By.CLASS_NAME, "block-iframe")
        videolist = []
        for video in videos:
            HTML = video.get_attribute('outerHTML')        
            videolist.append(re.search(r'src=(.+?)"',HTML)[0].replace('src="','').replace('"',''))
            if len(videolist)!=0:
                dictQuestion['video_lesson'] = [videolist]
    except:
        "without_video_lesson"
    
def getlistID():
    """INITIAL FUNTION TO GET ALL THE ID VALUES THAT CONTAIN ALL THE BLOCK QUESTION"""    
    cardlist = driver.find_elements(By.CLASS_NAME, "card.card--nnki-question.ng-star-inserted")
    IDquestionlist = []
    for cardcontent in cardlist:
        questionHTML= cardcontent.get_attribute('outerHTML')
        idquestion = re.findall(r'question-\d+', questionHTML)        
        IDquestionlist.extend(idquestion)
    return IDquestionlist

def getItemnumber(question):
    """FUNCTION TO GET ITEM NUMBER, IT MAKE CLICK IN SECTION "'Strumenti per il ripasso'" AND GET THE CORRESPONDING ITEM NUMBER"""
    numbertry = 0    
    itemnumberflag = False
    UP = False
    while not itemnumberflag:
        try:
            # webdriver.ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            # click on "Strumenti per il ripasso" icon to get item number
            itemnumbers = question.find_element(By.CLASS_NAME, "ml-auto.mr-8.cursor-pointer.flex.items-center.text-12.space-x-2.ng-star-inserted")    
            itemnumbers.click()    
            itemnumberflag = True

        except:
            timewait = random.uniform(2, 8)
            time.sleep(timewait)
            numbertry +=1
            if numbertry>=10:
                webdriver.ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
                confirm = input('Confirm to continue without item number ')
                numbertry = 0
                if confirm =='y':
                    itemnumberflag = True

    maxtry = False
    numbertry = 0
    itemnumber = 'Empty Number'
    while not maxtry and itemnumberflag:
        try:
            itemnumber = driver.find_element(By.CLASS_NAME,"mat-tab-label-content").text.replace('ITEM ','')            
            if itemnumber !='':
                maxtry = True
                timewait = random.uniform(1, 5)
                time.sleep(timewait)
                print("wait itemnumber")
        except:
            print("New waiting")
            # if UP:
            if maxtry > 2:
                webdriver.ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
                time.sleep(0.5)
            # UP = False
            # else:
            #     webdriver.ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            #     time.sleep(0.5)
            #     UP = True
            timewait = random.uniform(0.5, 3)
            time.sleep(timewait)
            numbertry +=1
            if numbertry == 10:
                itemnumber = 'Empty Number'
                confirmation = input("\n Type 'y' to confirm that item numbert don't exist ")
                if confirmation =='y':
                    maxtry = True
    
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(2)
    return itemnumber

def softColumns():    
    print("# Original listcolumns #")
    print(listcolumns)
    listcorrectionimages = getlist('CorrectionImage_')    
    print("listcorrectionimages", listcorrectionimages)
    
    listcorrectionlinks = getlist('CorrectionLink_','href')
    print('listcorrectionlinks', listcorrectionlinks)
    
    listimagelinks = getlist('QuestionImage_src')
    
    
    baselist = ['Domanda','item','Question']
    baselist.extend(listimagelinks)
    baselist.extend(['A', 'B', 'C', 'D', 'E', 'CorrectOption','Correction'])
    baselist.extend(listcorrectionimages)
    baselist.extend(listcorrectionlinks)
    
    if 'video_lesson' in list(dictQuestion.keys()):
        baselist.append('video_lesson')

    print("Output columns list")
    print(baselist)
    return baselist

def getlist(word, word2=''):
    global listcolumns    
    listmatch = []
    for column in listcolumns:
        if word2 != '':
            if word in column and word2 in column and not(column in listmatch):
                listmatch.append(column)
        else:
            if word in column and not(column in listmatch):
                listmatch.append(column)
    return listmatch

def getExamInfo(filename, examlink):
    global listcolumns, numberquestion, dictQuestion
    df_all = pd.DataFrame()    
    driver.get(examlink)
    time.sleep(4)
    startStoptest(examlink)
        
    webdriver.ActionChains(driver).send_keys(Keys.PAGE_DOWN).perform()
    time.sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.PAGE_UP).perform()

    # FLAG VARIABLE TO INDICATE IF VIDEO LINK APPEAR AT BEGGINING WAS LOADED
    flagvideolesson = True
    
    IDquestionlist = getlistID()
    numberquestion = len(IDquestionlist)

    print("Number of question: ", numberquestion)
    for idquestion in IDquestionlist:
        dictQuestion = {}
        question = driver.find_element(By.ID, idquestion)

        QuestionID = getQuestionNumber(question)
        if 'Strumenti per il ripasso' in question.text:
            dictQuestion['item'] = getItemnumber(question)
        else:
            dictQuestion['item'] = "Without item number"
        print(QuestionID, end=' ')
        getQuestionDescription(question)

        getOptions(question)
        getCorrectOption(question)
        getQuestionLinkImage(question)
        
        answerexplenation = getCorrection(question)

        getLinkAnswer(answerexplenation)
        getLinkImageAnswer(answerexplenation)            
        # video lesson
        if flagvideolesson:
            flagvideolesson = False                
            getVideosLinks()
        print("Dict question before save: ")
        print(dictQuestion)
        ########################################
        #           CHECK QUESTION CONTENT     #
        ########################################
        print("#"*50, "LEN OF QUESTION DICT" ,"#"*50)
        print("#"*50, len(dictQuestion['Question']), "#"*50)
        if dictQuestion['Question']=='':
            confirm = input("Check question content: ")
            if confirm =='y':
                getExamInfo(filename, examlink)
            else:
                print('Continue processing: ')
        ########################################

        df = pd.DataFrame.from_dict(dictQuestion)            
        df_all = pd.concat([df_all, df])                
    
    df_all = df_all.reset_index(drop= True)
    sotflist = softColumns()        
    df_all = df_all[sotflist]

    df_all = df_all.fillna('*')
    df_all.to_excel(filename + ".xlsx")  
    return True
#########################################################################################
#                                                                                       #
#                   SECTION FOR STAR TEST AND THEN STOP                                 #
#                                                                                       #
#########################################################################################
def startStoptest(examlink):

    try:        
        clickCorrectionButton()        
    except:
        
        try:
            dictverification = {}
            clickStartButton() #Click  button start button            
            dictverification['START']= 'READY'
            confirmStart() # Click in emerging windows to start test.            
            dictverification['CONFIRM_START']= 'READY'
            
            print("Next step click in stopbutton , try to find stop button")
            flagstopbutton = clickStopButton() # Return True if found button stop
            print("FLAG STOP: ", flagstopbutton)
            
            if not flagstopbutton:
                print("Load previous page")
                driver.get(examlink)# Load previus page or exam link                
                driver.switch_to.alert.accept()                
                dictverification['LOAD PREVIOUS PAGE']= 'READY'
                time.sleep(random.uniform(3, 8))            
        except:
            print("Not emerging windows")
            print(dictverification)
            
        print("Click on emerged windows to continue")
        
        if flagstopbutton:
            # Click on share button
            clickSharebuton()        
        clickCorrectionButton()
        # Click in corrections butto
        
def clickCorrectionButton():
    correccionesbutton = driver.find_element(By.CLASS_NAME, 'mat-tooltip-trigger.cursor-pointer.text-site-main-one.font-semibold.ng-star-inserted')
    correccionesbutton.click()    
    time.sleep(random.uniform(3, 8))
    # print("Sent escape key to close aid")
    # webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform() # generar ESC para salir de la conversacion

def clickStartButton():
    buttonstart = driver.find_element(By.CLASS_NAME, 'mat-ripple.btn-nnki.btn-nnki-primary.btn-nnki-uppercase.btn-nnki-heavy.mat-ripple-unbounded.ng-star-inserted')
    buttonstart.click()
    print("Cliked in button start test")    
    time.sleep(random.uniform(3, 8))

def confirmStart():
    try:
        confirmstart = driver.find_element(By.CLASS_NAME,'modal-buttons')    
        confirmstart.click()
        print("Click on emerging windows to confirm start test")        
        time.sleep(random.uniform(3, 8))
    except:
        print("Not emerging windows")

def serchStopButton(classename):
    # global dictverification

    stopbuttonblocks = driver.find_elements(By.CLASS_NAME, classename)

    for stopbuttonblock in stopbuttonblocks:
        if 'stop.svg'in stopbuttonblock.get_attribute('outerHTML'):        
            stopbuttonblock.click()            
            time.sleep(random.uniform(3, 8))
            print("Stop button clicked")
            # dictverification['STOP BUTTON CLICKED']= 'READY'            
            flagstopbutton = True
        else:
            flagstopbutton = False
    return flagstopbutton

def clickStopButton():
    global dictverification
    try:
        try:
            # stopbuttonblocks = driver.find_elements(By.CLASS_NAME, 'widget-button')
            flagstopbutton = serchStopButton('widget-button')
            print("flagstopbutton in general function option 1", flagstopbutton)
            # stopbuttonblocks = driver.find_elements(By.CLASS_NAME, 'widget-buttons.mr-auto')

        except:
            # stopbuttonblocks = driver.find_elements(By.CLASS_NAME, 'widget-buttons.mr-auto')
            flagstopbutton = serchStopButton('widget-buttons.mr-auto')
            print("flagstopbutton in general function option 2", flagstopbutton)
        # confirm stop test in emerging windows.
        print("Nex step find button to confirmate end test")
        confirmendtest = driver.find_element(By.CLASS_NAME,'mat-ripple.btn-nnki.btn-nnki-red.btn-nnki-uppercase.mat-ripple-unbounded')
        confirmendtest.click()
        print("Click on confirm end button")        
        time.sleep(random.uniform(3, 9))        
        return True
    except:
        return False        

def clickSharebuton():    
    # Click in button share
    sharebutton = driver.find_element(By.CLASS_NAME, 'modal-buttons')#'mat-ripple.btn-nnki.btn-nnki-primary.btn-nnki-uppercase.mat-ripple-unbounded.ng-star-inserted')
    sharebutton.click()    
    time.sleep(random.uniform(3, 8))
    print("Click in button share")

def waitOfferWindows():
    windowsoffer = False
    countlimit = 0
    while not windowsoffer:
        try:
            offerbutton = driver.find_element(By.CLASS_NAME, 'mat-ripple.btn-nnki.btn-nnki-black.btn-nnki-uppercase.mat-ripple-unbounded')
            if offerbutton.text!="":            
                windowsoffer = True
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except:            
            time.sleep(1)
            countlimit +=1
        if countlimit==10:
            windowsoffer = True

#########################################################################################
#                                                                                       #
#                   SECTION GET ALL LINKS AVAILABLES                                    #
#                                                                                       #
#########################################################################################
def getExamsLinks():
    input("Remember make zoom out and scrowll down until load all the tests,\n Then Input 'y' to continue ")
    listExams = driver.find_elements(By.CLASS_NAME,"card-header")
    dictlinks = {}
    for n, examlink in enumerate(listExams):
        link = examlink.get_attribute('href')
        title = examlink.find_element(By.CLASS_NAME, 'card-title').text.replace(' ','_')
        topic = examlink.find_element(By.CLASS_NAME, 'card-theme').text.replace(' ','_')        
        link = examlink.get_attribute('href')        
        dictlinks[n] = {'title':title,'topic':topic,'link':link}                
    saveCheckPoint('files_prepssm/examslinks.json', dictlinks)

#########################################################################################
#                                                                                       #
#                   SECTION TO CONTROL EXAMS DOWNLOAD                                   #
#                                                                                       #
#########################################################################################
def keysPending():
    """The funcion check files timeexecution.json and examslinks.json
    and get the list of numbers of the pending links """
    dictverification = loadCheckPoint('files_prepssm/timesexecution.json')
    keysready = list(dictverification.keys())
    print(keysready,'\n')

    dictlinks = loadCheckPoint('files_prepssm/examslinks.json')
    completekeys = list(dictlinks.keys())

    for keyready in keysready:    
        completekeys.remove(keyready)
    return completekeys

def getIndexPendingFiles():
    examslinks = loadCheckPoint("files_prepssm/examslinks.json")
    listallfiles = []
    
    for key in examslinks.keys():        
        listallfiles.append(examslinks[key]['title']+'.xlsx')
    
    path = 'files_prepssm/' 
    files = os.listdir(path)

    files_txt = [i for i in files if i.endswith('.xlsx')]

    listpendingindex = list(examslinks.keys())    
    indexready = []
    for filename in files_txt:
        indexfile = listallfiles.index(filename)        
        indexready.append(indexfile)
        listpendingindex.remove(str(indexfile))        
        
    print("INDEX READY: ", sorted(indexready) )
    print("Total files ready: ", len(indexready))
    return listpendingindex
#########################################################################################
#                                                                                       #
#                   MAIN LOOP ITERATE OVER ALL THE LINKS                                #
#                                                                                       #
#########################################################################################
def loopOverLinks(dictlinks):
    global listcolumns
    # try:
    # userconfirmation = input("Do you want to continue in previous check point? 'y' yes or 'n' to restart at inicial point ")
    userconfirmation = 'y'
    if userconfirmation == 'y':
        checkpoint = loadCheckPoint('files_prepssm/lastexam.json')
        idnumber = int(checkpoint['last'])
        idnumber += 1
        print("Loading check point, last dict number: ", idnumber)
    else:
        idnumber = 0

    # fileslinks = {}
    
    # for keynumber, examLink in enumerate(listprovisional):        
    # for keynumber in list(dictlinks.keys())[0:5][idnumber:]:
    repeatprocess = True
    totaltime = 0
    while repeatprocess:
        timesdict = loadCheckPoint('files_prepssm/timesexecution.json')
        # keys_pending_dowload = keysPending()
        # keys_pending_dowload = getIndexPendingFiles()
        # keys_pending_dowload.remove('51')# GUIDE TO UTILIZE
        
        # keys_pending_dowload = [0, 5, 6, 8, 9, 10, 14, 22, 27,34, 74, 78, 115,117,120]
        keys_pending_dowload = [285]
        print("\n Pending index: ", keys_pending_dowload)
        try:
            for keynumber in keys_pending_dowload:

                start = time.time()
                examLink = dictlinks[str(keynumber)]['link']                
                name = dictlinks[str(keynumber)]['title']
                print("Actual keynumber", keynumber)
                print("file name", name)
                print("Link: ", examLink)                
                filename = 'files_prepssm/' + name
                # fileslinks[keynumber] ={'file':filename.split('/')[-1],'link':examLink}# To save file name and link, # DELETE #
                listcolumns = ['item','Question','QuestionImage_src','A', 'B', 'C', 'D', 'E', 'CorrectOption','Correction']
                flagsuccess = getExamInfo(filename, examLink)
                saveLastArticle(keynumber, examLink)
                # input("Continue no next link ")
                # save json file to registe time in each file dowload.
                dictexamslinks = loadCheckPoint('files_prepssm/correctionliks.json')
                dictexamslinks[keynumber] = driver.current_url
                saveCheckPoint('files_prepssm/correctionliks.json', dictexamslinks)
                end = time.time()
                timexcution = end - start            
                #####################################################################
                #   BLOCK TO SAVE REGISTER TIME AT SUCCESS DOWNLOADS FILES          #
                #                                                                   #
                #####################################################################

                if flagsuccess:
                    executiontimes = {"numbQuestion":numberquestion, 'time':timexcution}
                    timesdict[keynumber] = executiontimes
                    saveCheckPoint('files_prepssm/timesexecution.json', timesdict)
                timewait = random.uniform(5, 30)
                time.sleep(timewait)
                #####################################################################
                #   GENERATE LONG TIME SLEEP WHEN PASS OF 5 HOURS                   #
                #                                                                   #
                #####################################################################
                totaltime = totaltime + timexcution
                # if totaltime >= 18000:
                #     timewait = random.uniform(900, 2000)
                #     time.sleep(timewait)
                #     print("Time to take a snach", timewait)
                #     totaltime = 0
        # saveCheckPoint('files_prepssm/fileslinks.json', fileslinks)
        # print(fileslinks)

        except Exception as e:
            print("Error: ")
            print(e)
            # executiontimes = {"numbQuestion":numberquestion, 'time':"ERROR"}
            # timesdict[keynumber] = executiontimes
            # saveCheckPoint('files_prepssm/timesexecution.json', timesdict)

            answer = input("Confirm error verification, Do you want repeat the process?, Type 'y' to continue other to cancel ")
            if answer != 'y':
                repeatprocess = False
    print("Loop over links finished ")
    #     print("Issue present in Loop over links")
    #     print("#"*80)
    #     print(e)
    #     input("Input 'y' to continue ")

#########################################################################################
#                                                                                       #
#                   SECTION TO COMPLETE MISSING DATA                                    #
#                                                                                       #
#########################################################################################
def getDictIssues():
    ############################# GET ONLY FILES NAMES LIST #############################
    examslinks = loadCheckPoint("files_prepssm/examslinks.json")
    listallfiles = []    
    for key in examslinks.keys():        
        listallfiles.append(examslinks[key]['title']+'.xlsx')

    ############################# LOOP OVER LIST OF FILES #############################
    path = 'files_prepssm/' 
    files = os.listdir(path) # LOAD THE LIST OF FILES WITH TERMINATION .xlsx

    files_xlsx = [i for i in files if i.endswith('.xlsx')]

    dictIssues = {} # CREATE DICT TO SAVE ISSUES FILES

    # LIST OF CONDITIONS

    for filename in files_xlsx:

        df = pd.read_excel('files_prepssm/'+filename)
        cond1 = len(df[df['item']=='Empty Number'])!=0 
        cond2 = len(df[df['item']=='Without item number'])!=0 
        cond3 = len(df[df['item']==''])!=0

        if cond1 or cond2 or cond3:

            indexfile = listallfiles.index(filename)
            dictIssues[indexfile] = {'title':examslinks[str(indexfile)]['title'],
                                     'link':examslinks[str(indexfile)]['link'], 
                                     'FIXED':'False'
                                     }
    saveCheckPoint('files_prepssm/dictIssues.json', dictIssues)
    return dictIssues

def iterateOnDictIssues():

    # listIssues = sorted(list(dictIssues.keys()))
    dictIssues = loadCheckPoint('files_prepssm/dictIssues.json')
    
    for KEY in dictIssues.keys():
        if dictIssues[KEY]['FIXED'] =='False':
            currentIndex = KEY
            break
    print(type(currentIndex))
    if dictIssues[currentIndex]['FIXED']=='False':
        print("Current Index: ", currentIndex)
        print("File Title: ", dictIssues[currentIndex]['title'])
        examLink = dictIssues[currentIndex]['link']
        print("Link: ", examLink )
        driver.get(dictIssues[currentIndex]['link'])
        time.sleep(4)
        startStoptest(examlink)
        filename = 'files_prepssm/' + dictIssues[currentIndex]['title']+'.xlsx'
        df = pd.read_excel(filename)
        cond1 =  df['item']=='Without item number' 
        cond2 = df['item']=='Empty Number'
        cond3 = df['item']==''

    return df, df[cond1|cond2|cond3], currentIndex, filename, examlink

def getIndexRepeat():
    """ Function to get index of files that required repeat """
    dictIssues = loadCheckPoint('files_prepssm/dictIssues.json')
    listRepeat = []
    for KEY in dictIssues.keys():
        if dictIssues[KEY]['FIXED'] =='Repeat':
            listRepeat.append(KEY)           
    print(listRepeat)
    
def getMissingData(df_filtered, examlink):
    global dictQuestion
    listindex = df_filtered.index.values    

    webdriver.ActionChains(driver).send_keys(Keys.PAGE_DOWN).perform()
    time.sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.PAGE_UP).perform()

    IDquestionlist = getlistID()
    numberquestion = len(IDquestionlist)

    print("Number of question: ", numberquestion)
    dict_to_complete = {}
    # for idquestion in IDquestionlist:
    for index in listindex:
        idquestion = IDquestionlist[index]
        dictQuestion = {}
        question = driver.find_element(By.ID, idquestion)

        QuestionID = getQuestionNumber(question)
        if 'Strumenti per il ripasso' in question.text:
            dictQuestion['item'] = getItemnumber(question)
        else:
            dictQuestion['item'] = "Without item number"
        print(QuestionID, end=' ')
        getQuestionDescription(question)

        print("dictQuestion['Question']", dictQuestion['Question'])
        print("df_filtered['Question'][index]", df_filtered['Question'][index])
        
        if dictQuestion['Question'][0] == df_filtered['Question'][index]:
            print("Confirmation Match content Question ")
            dict_to_complete[index] = dictQuestion['item']

        print("Resulted dict to complete: ", dict_to_complete)
    
    return dict_to_complete

def completeDataFrame(df, dict_to_complete):    

    for index in dict_to_complete.keys():       
        print("INDEX: ",index)
        df.at[index,'item']= dict_to_complete[index]
        
    return df

def confirmFileRevision(currentIndex):
    dictIssues = loadCheckPoint('files_prepssm/dictIssues.json')
    print(dictIssues[currentIndex]["title"])
    print(dictIssues[currentIndex]["link"])
    confirm = input('Do you confirm the file was checked? , type "yes" ')
    if confirm =='yes':

        dictIssues[currentIndex]["FIXED"] = 'True'        
        print("The list of files was updated succesfully")
        saveCheckPoint('files_prepssm/dictIssues.json', dictIssues)
    else:
        confirm = input('Current file is necessary repeat? , type "yes" ')
        if confirm =='yes':
            dictIssues[currentIndex]["FIXED"] = 'Repeat'            
            print("The list of files was updated succesfully REPEAT FILE")
            saveCheckPoint('files_prepssm/dictIssues.json', dictIssues)

def replaceFile(filename, df_completed, currentIndex):
    dictIssues = loadCheckPoint('files_prepssm/dictIssues.json')
    confirm = input('Do you are sure to replace file: , type "yes"')
    if confirm == 'yes':
        df_completed.to_excel(filename)
        print("File replaced")
        dictIssues[currentIndex]['FIXED'] = 'True'    
        saveCheckPoint('files_prepssm/dictIssues.json', dictIssues)

#########################################################################################
#                                                                                       #
#                   SECTION TO GENERATE REPORT                                          #
#                                                                                       #
#########################################################################################

def generateReportItem(reportname = 'deliverfiles/Item_report.xlsx'):
    if not os.path.isdir('deliverfiles'):
        os.mkdir('deliverfiles')

    df_all = pd.DataFrame()
    for KEY in dictExamsLinks.keys():
        dict1 = dictExamsLinks[KEY]

        dict1['topic'] = dict1['topic'].replace('_', ' ')
        dict1['title'] = dict1['title'].replace('_', ' ')
        df = pd.DataFrame(dict1, index=[0])
        df_all = pd.concat([df_all, df])    

    df_all = df_all[['title', 'topic', 'numbQuestion']]
    df_all = df_all.rename(columns={"title": "Item", "topic": "Name Item"})
    df_all = df_all.reset_index(drop=True)
    df_all.to_excel(reportname)

def concatItemFiles(dictExamsLinks, pathfiles):
    """ 

     """
    global listcolumns
    df_all = pd.DataFrame()
    for KEY in dictExamsLinks.keys():
        if KEY in ['108', '132', '146']:
            title = dictExamsLinks[KEY]['title'].split(':')[0]
            topic = dictExamsLinks[KEY]['topic'].split(':')[0]
        else:
            title = dictExamsLinks[KEY]['title']
            topic = dictExamsLinks[KEY]['topic']

        filename = pathfiles + clearfilename(title)+'_'+ clearfilename(topic)+'.xlsx'
    #         input("continue: ")
        # CHECK IF FILE EXIST
        cond1 = os.path.isfile(filename)

        if cond1:
            print('-',end='')
            df = pd.read_excel(filename)
            df_all = pd.concat([df_all, df])
        else:
            print("Unexist file: ", filename)
    listcolumns = df_all.columns

    print("listcolumns inside function: ", listcolumns)

    sofcolumns = softColumnsCompleteFile()
    df_all = df_all[sofcolumns].reset_index(drop = True)
    print("Total number of question: ",len(df_all))
    df_all.to_excel('deliverfiles/Item_complete.xlsx')

def softColumnsCompleteFile(df_columns):
    
    listcolumns = df_columns
    listcorrectionimages = getlist('CorrectionImage_')    
    print("listcorrectionimages", listcorrectionimages)
    
    listcorrectionlinks = getlist('CorrectionLink_','href')
    print('listcorrectionlinks', listcorrectionlinks)
    
    listimagelinks = getlist('QuestionImage_src')
    
    
    baselist = ['Domanda','item','Question']
    baselist.extend(listimagelinks)
    baselist.extend(['A', 'B', 'C', 'D', 'E', 'CorrectOption','Correction'])
    baselist.extend(listcorrectionimages)
    baselist.extend(listcorrectionlinks)
    
    if 'video_lesson' in list(dictQuestion.keys()):
        baselist.append('video_lesson')

    print("Output columns list")
    print(baselist)
    return baselist

chromedriver_autoinstaller.install() 
 
# Create Chromeoptions instance 
options = webdriver.ChromeOptions() 
 
# Adding argument to disable the AutomationControlled flag 
options.add_argument("--disable-blink-features=AutomationControlled") 
 
# Exclude the collection of enable-automation switches 
options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
 
# Turn-off userAutomationExtension 
options.add_experimental_option("useAutomationExtension", False)
 
# Changing the property of the navigator value for webdriver to undefined 
# driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") 


#####################################################################
#                                                                   #
#               Profile chrome configuration.                       #
#                                                                   #
#####################################################################

# Define default profiles folder
options.add_argument(r"user-data-dir=/home/jorge/.config/google-chrome/")
# Define profile folder, profile number
options.add_argument(r"profile-directory=Profile 6")

# global s
driver = webdriver.Chrome(options=options)
dictQuestion = {}
listcolumns = ['item','Question','QuestionImage_src','A', 'B', 'C', 'D', 'E', 'CorrectOption','Correction']
numberquestion = 0
def __init__():  
    global listcolumns

    confirm = input("Confirm if login is ready, type 'y' to continue ")
    while not(confirm!='y' or confirm!='n'):
        confirm = input("Please type a valid option 'y' or 'n' ")

    
    if confirm=='y':

        driver.get("https://www.prepssm.it/app/exams/examens")

        # loadprevious = input("Do you want to load previous list? Type'y' load file and 'n' upload list of exam ")
        loadprevious = 'y'
        if loadprevious=='n':
            # Upload or get all exam links again
            # driver.get("https://www.prepssm.it/app/exams/examens")
            getExamsLinks()
            loadprevious ='y'
        
        if loadprevious=='y':
            dictlinks = loadCheckPoint('files_prepssm/examslinks.json')
            loopOverLinks(dictlinks)
            # Loop on list of links and load previous checkpoint.

        time.sleep(1500)
        input("Type anything to close")
    else:
        print("Script canceled ")

__init__()