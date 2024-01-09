#import cv2
import os
import cv2
from natsort import natsorted,natsort_keygen
import openpyxl
#import xml2json.xml2jsonn
import xmltodict
#import xml2jsonn
import pandas as pd
import uuid
import json
#file = xmltodict.parse(r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_dêp\slide1.xml')
from xml.dom import minidom
# parse file items.xml
import codecs
import xml.etree.ElementTree as ET
import elementree_withedittext#.lay_all_objtext_infilexml
types_of_encoding = ["utf8", "cp1252"]
from xml.dom.minidom import parse, parseString
#import xml2json.xml2jsonn as xml2jsoni
import edit_colon2_
#method load all link xml in folder

link_folder_saveText="./text"

def load_filetext_from_folder(folder):
    XMLS = []
    link=[]
    newfoler=[]
    listtitle=[]
    #a=r

    for filename in os.listdir(folder + str(r"""\text""")):
        #print(filename)
        #xml = cv2.imread(os.path.join(folder + """\ppt""" + """\slides""", filename))
        #xml = cv2.imread(os.path.join(folder, filename))
        xml = filename
        #xml = natsorted(os.listdir(folder + """\ppt""" + """\slides"""))

        if xml is not None:
            XMLS.append(xml)
            #link.append(folder + """\ppt""" + """\slides\%s"""%(xml))
            listtitle.append(xml[:-4])
        #XMLS.sort(key=lambda x: float(x.strip('slide')))
        #XMLS.remove('_rels')
        #XMLS.append(xml)
    #natsort_key = natsort_keygen()
    #print(natsort_key)
    #newxml=natsorted(XMLS)
    newxml = list.copy(natsorted(XMLS))
    newlisttitle= list.copy(natsorted(listtitle))
    for filename in XMLS:
        #link.append(os.listdir(folder str( """\ppt""") + str("""\slides"""+r"""\%s"""%(filename))) )
        link.append(os.path.join(folder + str(r"""\text"""), filename))
        #newfoler.append()
    newlink = list.copy(natsorted(link))

    #xmlsss=newxml[]
    #link_xml_cuapptx = folder + """\ppt""" + """\slides"""
    # for i in newxml:
    #     print(i)
    # for i in newlink:
    #     print(i)
    # for i in newlisttitle:
    #     print(i)
    return newxml,newlink,newlisttitle#,newfoler
def load_filetext_from_folder1(folder):
    XMLS = []
    link=[]
    newfoler=[]
    listtitle=[]
    #a=r

    for filename in os.listdir(folder):
        #print(filename)
        #xml = cv2.imread(os.path.join(folder + """\ppt""" + """\slides""", filename))
        #xml = cv2.imread(os.path.join(folder, filename))
        xml = filename
        #xml = natsorted(os.listdir(folder + """\ppt""" + """\slides"""))

        if xml is not None:
            XMLS.append(xml)
            #link.append(folder + """\ppt""" + """\slides\%s"""%(xml))
            listtitle.append(xml[:-4])
        #XMLS.sort(key=lambda x: float(x.strip('slide')))
        #XMLS.remove('_rels')
        #XMLS.append(xml)
    #natsort_key = natsort_keygen()
    #print(natsort_key)
    #newxml=natsorted(XMLS)
    newxml = list.copy(natsorted(XMLS))
    newlisttitle= list.copy(natsorted(listtitle))
    for filename in XMLS:
        #link.append(os.listdir(folder str( """\ppt""") + str("""\slides"""+r"""\%s"""%(filename))) )
        link.append(os.path.join(folder), filename)
        #newfoler.append()
    newlink = list.copy(natsorted(link))

    #xmlsss=newxml[]
    #link_xml_cuapptx = folder + """\ppt""" + """\slides"""
    # for i in newxml:
    #     print(i)
    # for i in newlink:
    #     print(i)
    # for i in newlisttitle:
    #     print(i)
    return newxml,newlink,newlisttitle#,newfoler
def read_text_from_file_text(link_file_text):
    list_quiz=[]
    Array_quest=[]
    with open(link_file_text, encoding='utf-8') as f:

        for line in f:
            #print(line.strip())
            Array_quest.append(str(line.strip()))
    return Array_quest
def read_text_from_file_text_xlsx(link_file_text):


    # Define variable to load the dataframe


    # read by default 1st sheet of an excel file
    dataframe1 = pd.read_excel(link_file_text)

    print(dataframe1)
def add_h5p_json_quest(arrray,type_quizz):
    #str(nhap4.uuid.uuid4())



    if type_quizz == 'quizz':
        h5p_Quiz =  {
            "x": 27.233115468409586,
            "y": 23.733135988711656,
            "width": 50,
            "height": 50,
            "action": {
              "library": "H5P.SingleChoiceSet 1.11",
              "params": {
                "choices": [
                  {
                    "subContentId": str(uuid.uuid4()),
                    "question": "<p>%s</p>\n" % arrray[0],
                    "answers": [
                      "<p>%s</p>\n"%arrray[1],
                      "<p>%s</p>\n"%arrray[2],
                      "<p>%s</p>\n"%arrray[3],
                      "<p>%s</p>\n"%arrray[4]
                    ]
                  },
                  {
                    "subContentId": str(uuid.uuid4())
                  }
                ],
                "overallFeedback": [
                  {
                    "from": 0,
                    "to": 100
                  }
                ],
                "behaviour": {
                  "autoContinue": 1,
                  "timeoutCorrect": 2000,
                  "timeoutWrong": 3000,
                  "soundEffectsEnabled": 1,
                  "enableRetry": 1,
                  "enableSolutionsButton": 1,
                  "passPercentage": 100
                },
                "l10n": {
                  "nextButtonLabel": "Next question",
                  "showSolutionButtonLabel": "Show solution",
                  "retryButtonLabel": "Retry",
                  "solutionViewTitle": "Solution list",
                  "correctText": "Correct!",
                  "incorrectText": "Incorrect!",
                  "muteButtonLabel": "Mute feedback sound",
                  "closeButtonLabel": "Close",
                  "slideOfTotal": "Slide :num of :total",
                  "scoreBarLabel": "You got :num out of :total points",
                  "solutionListQuestionNumber": "Question :num",
                  "a11yShowSolution": "Show the solution. The task will be marked with its correct solution.",
                  "a11yRetry": "Retry the task. Reset all responses and start the task over again."
                }
              },
              "subContentId": str(uuid.uuid4()),
              "metadata": {
                "contentType": "Single Choice Set",
                "license": "U",
                "title": "Untitled Single Choice Set",
                "authors": [],
                "changes": [],
                "extraTitle": "Untitled Single Choice Set"
              }
            },
            "alwaysDisplayComments": 0,
            "backgroundOpacity": 0,
            "displayAsButton": 0,
            "buttonSize": "big",
            "goToSlideType": "specified",
            "invisible": 0,
            "solution": ""
          }
        return h5p_Quiz
    if type_quizz == 'multil_choice':
        h5p_multil_choice={
              "x": 18.51851851851852,
              "y": 28.048251623022868,
              "width": 59.91143790849673,
              "height": 51.78138761173452,
              "action": {
                "library": "H5P.MultiChoice 1.16",
                "params": {
                  "media": {
                    "disableImageZooming": 0
                  },
                  "answers": [
                    {
                      "correct": 1,
                      "tipsAndFeedback": {
                        "tip": "",
                        "chosenFeedback": "",
                        "notChosenFeedback": ""
                      },
                      "text": "<div>%s</div>\n" % arrray[5]
                    },
                    {
                      "correct": 0,
                      "tipsAndFeedback": {
                        "tip": "",
                        "chosenFeedback": "",
                        "notChosenFeedback": ""
                      },
                      "text": "<div>%s</div>\n"%arrray[2]
                    },
                    {
                      "correct": 0,
                      "tipsAndFeedback": {
                        "tip": "",
                        "chosenFeedback": "",
                        "notChosenFeedback": ""
                      },
                      "text": "<div>%s</div>\n"%arrray[1]
                    }
                  ],
                  "overallFeedback": [
                    {
                      "from": 0,
                      "to": 100
                    }
                  ],
                  "behaviour": {
                    "enableRetry": 1,
                    "enableSolutionsButton": 1,
                    "enableCheckButton": 1,
                    "type": "auto",
                    "singlePoint": 0,
                    "randomAnswers": 1,
                    "showSolutionsRequiresInput": 1,
                    "confirmCheckDialog": 0,
                    "confirmRetryDialog": 0,
                    "autoCheck": 0,
                    "passPercentage": 100,
                    "showScorePoints": 1
                  },
                  "UI": {
                    "checkAnswerButton": "Check",
                    "submitAnswerButton": "Submit",
                    "showSolutionButton": "Show solution",
                    "tryAgainButton": "Retry",
                    "tipsLabel": "Show tip",
                    "scoreBarLabel": "You got :num out of :total points",
                    "tipAvailable": "Tip available",
                    "feedbackAvailable": "Feedback available",
                    "readFeedback": "Read feedback",
                    "wrongAnswer": "Wrong answer",
                    "correctAnswer": "Correct answer",
                    "shouldCheck": "Should have been checked",
                    "shouldNotCheck": "Should not have been checked",
                    "noInput": "Please answer before viewing the solution",
                    "a11yCheck": "Check the answers. The responses will be marked as correct, incorrect, or unanswered.",
                    "a11yShowSolution": "Show the solution. The task will be marked with its correct solution.",
                    "a11yRetry": "Retry the task. Reset all responses and start the task over again."
                  },
                  "confirmCheck": {
                    "header": "Finish ?",
                    "body": "Are you sure you wish to finish ?",
                    "cancelLabel": "Cancel",
                    "confirmLabel": "Finish"
                  },
                  "confirmRetry": {
                    "header": "Retry ?",
                    "body": "Are you sure you wish to retry ?",
                    "cancelLabel": "Cancel",
                    "confirmLabel": "Confirm"
                  },
                  "question": "<p>%s</p>\n"%arrray[0]
                },
                "subContentId": str(uuid.uuid4()),
                "metadata": {
                  "contentType": "Multiple Choice",
                  "license": "U",
                  "title": "Untitled Multiple Choice",
                  "authors": [],
                  "changes": [],
                  "extraTitle": "Untitled Multiple Choice"
                }
              },
              "alwaysDisplayComments": 0,
              "backgroundOpacity": 0,
              "displayAsButton": 0,
              "buttonSize": "big",
              "goToSlideType": "specified",
              "invisible": 0,
              "solution": ""
            }
        return h5p_multil_choice
    if type_quizz == 'true_false':
        H5p_true_false= {
            "x": 30,
            "y": 30.00084144754869,
            "width": 40,
            "height": 40,
            "action": {
              "library": "H5P.10 1.8",
              "params": {
                "media": {
                  "disableImageZooming": 0
                },
                "correct": "0",
                "behaviour": {
                  "enableRetry": 1,
                  "enableSolutionsButton": 1,
                  "enableCheckButton": 1,
                  "confirmCheckDialog": 0,
                  "confirmRetryDialog": 0,
                  "autoCheck": 0
                },
                "l10n": {
                  "1Text": "True",
                  "0Text": "False",
                  "score": "You got @score of @total points",
                  "checkAnswer": "Check",
                  "submitAnswer": "Submit",
                  "showSolutionButton": "Show solution",
                  "tryAgain": "Retry",
                  "wrongAnswerMessage": "Wrong answer",
                  "correctAnswerMessage": "Correct answer",
                  "scoreBarLabel": "You got :num out of :total points",
                  "a11yCheck": "Check the answers. The responses will be marked as correct, incorrect, or unanswered.",
                  "a11yShowSolution": "Show the solution. The task will be marked with its correct solution.",
                  "a11yRetry": "Retry the task. Reset all responses and start the task over again."
                },
                "confirmCheck": {
                  "header": "Finish ?",
                  "body": "Are you sure you wish to finish ?",
                  "cancelLabel": "Cancel",
                  "confirmLabel": "Finish"
                },
                "confirmRetry": {
                  "header": "Retry ?",
                  "body": "Are you sure you wish to retry ?",
                  "cancelLabel": "Cancel",
                  "confirmLabel": "Confirm"
                },
                "question": "<p>%s</p>\n"%arrray[0]
              },
              "subContentId": str(uuid.uuid4()),
              "metadata": {
                "contentType": "1/0 Question",
                "license": "U",
                "title": "Untitled 1/0 Question",
                "authors": [],
                "changes": [],
                "extraTitle": "Untitled 1/0 Question"
              }
            },
            "alwaysDisplayComments": 0,
            "backgroundOpacity": 0,
            "displayAsButton": 0,
            "buttonSize": "big",
            "goToSlideType": "specified",
            "invisible": 0,
            "solution": ""
          }
        return H5p_true_false

if __name__ == "__main__":
  link1=r"search-ms:displayname=Search%20Results%20in%20pptx2h5p_1_2_deep_copy_addimages_NCKH&crumb=System.Generic.String%3Aget&crumb=location:C%3A%5CUsers%5CACER%5Canaconda3%5Cenvs%5Ch5p%5Cpptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001%5Cpptx2h5p_1_2_deep_copy_addimages_NCKH"
    # a,b,c=load_filetext_from_folder(r"C:\Users\ACER\anaconda3\envs\h5p\pptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001\pptx2h5p_1_2_deep_copy_addimages_NCKH")
    #"""vi du test link - """
    # for i in a:
    #     print(i)
    # for i in b:
    #     print(i)
    # for i in c:
    #     print(i)
  # """vi du add quizz"""
  # list_h5p_object_quest=[]
  # for i, link in enumerate(b):
  #     if "true_false" in link :
  #         print("true_false ok")
  #         az = read_text_from_file_text(link)
  #         #print("gia tri true false")
  #         print(add_h5p_json_quest(az,"true_false" ))
  #         list_h5p_object_quest.append(add_h5p_json_quest(az,"true_false" ))
  #     if "multil_choice" in link :
  #         print("multil_choice ok")

  #         az = read_text_from_file_text(link)
  #         print(add_h5p_json_quest(az, "multil_choice"))
  #         list_h5p_object_quest.append(add_h5p_json_quest(az,"multil_choice" ))
  #     if "quizz" in link:
  #         print("quizz ok")
  #         az = read_text_from_file_text(link)
  #         print(add_h5p_json_quest(az,"quizz"))
  #         list_h5p_object_quest.append(add_h5p_json_quest(az,"quizz" ))
  
  
  
  # for i in list_h5p_object_quest:
  #     print("gia tri i:",i)
  # print("gia tri b:",b)
  # link= r'C:\Users\ACER\anaconda3\envs\h5p\pptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001\pptx2h5p_1_2_deep_copy_addimages_NCKH\text\1.txt'

  # with open(r'C:\Users\ACER\anaconda3\envs\h5p\pptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001\pptx2h5p_1_2_deep_copy_addimages_NCKH\text\1.txt', encoding='utf8') as f:
  #     for line in f:
  #         print(line.strip())
  #         print('1')
  # read_text_from_file_text_xlsx(link)

  # for i in c:
  #     print(i)
  # read_text_from_file_text(b)
  a,b,c=load_filetext_from_folder1(link1)

  for i, link in enumerate(b):
          if ".py" in link :
              print("PY ok")
              print(link[-3:])