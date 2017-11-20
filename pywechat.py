# -*- coding: utf-8 -*-
"""
Created on Mon Nov 20 13:19:20 2017

@author: hfutzf
"""
from PIL import Image
import itchat
import win32com.client 
import cv2,os
import win32api,time
import numpy as np


voice= win32com.client.Dispatch("SAPI.SpVoice")
itchat.login()

itchat.send('python——微信宿舍监控系统\n1.输入cap:发送宿舍实时图像信息。\n2.输入video:发送宿\
舍实时视频监控信息\n3.发送@+文件（或者程序），远程打开文件\n4.发送\
除上述情况外的文本，远程实时文本转语音输出',toUserName='filehelper')

@itchat.msg_register('Text') #注册文本消息
def text_reply(msg): #心跳程序
    global flag
    message =  msg['Text'] #接收文本消息
    toName = msg['ToUserName'] #接收方

    if toName == "filehelper":
        if message == "cap": #远程拍照并发送到手机
        
            cap = cv2.VideoCapture(0)
            fourcc = cv2.VideoWriter_fourcc(*'XVID')
            out = cv2.VideoWriter('output.mp4',fourcc, 5.0, (640,480))
            
            start_time=time.clock()
            while(cap.isOpened() and (time.clock()-start_time)<3):
                ret, frame = cap.read()
                if ret==True:
                    frame = cv2.flip(frame,0)
                    out.write(frame)
                    
                    # 释放内存
            cap.release()
            out.release()
            
            vc = cv2.VideoCapture('output.mp4') #读入视频文件  
            c=1  
  
            if vc.isOpened(): #判断是否正常打开  
                rval , frame = vc.read()  
            else:  
                rval = False  
  
            timeF = 4  #视频帧计数间隔频率  
  
            while rval:   #循环读取视频帧  
                rval, frame = vc.read()  
                if(c%timeF == 0): #每隔timeF帧进行存储操作  
                      cv2.imwrite('image'+str(4) + '.jpg',frame) #存储为图像  
                c = c + 1  
                cv2.waitKey(1)  
            vc.release()  
            
            im1 = Image.open("image5.jpg")
            im2 = im1.rotate(180)
            im2.save('image.jpg')
            itchat.send('@img@%s'%u'image.jpg',toUserName='filehelper')
        if "@" not in message[0] or message !='cap' or message !='video' :
            voice.Speak(message)
        if message[0]=='@':
            if os.path.exists(message[1:]):
                try:
                    win32api.ShellExecute(0,'open', message[1:], '','',1)
                except:
                    pass
            else:
                itchat.send('打开文件失败',toUserName='filehelper')
        if message == "video":
          try:
            cap = cv2.VideoCapture(0)
            fourcc = cv2.VideoWriter_fourcc(*'XVID')
            out = cv2.VideoWriter('output.mp4',fourcc, 5.0, (640,480))
            
            start_time=time.clock()
            while(cap.isOpened() and (time.clock()-start_time)<5):
                ret, frame = cap.read()
                if ret==True:
                    frame = cv2.flip(frame,0)
                    out.write(frame)
                    
                    # 释放内存
            cap.release()
            out.release()
            itchat.send('@vid@%s'%u'output.mp4',toUserName='filehelper') 
          except:
              itchat.send('获取视频失败',toUserName='filehelper') 
itchat.run()

