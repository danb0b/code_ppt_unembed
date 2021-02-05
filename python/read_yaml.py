# -*- coding: utf-8 -*-
"""
Created on Thu Nov 21 16:35:18 2019

@author: danaukes
"""

import yaml
import subprocess
import os
#import sys
import shutil
import time
#import re
#import math
import glob
import video_tools.video_info as vi

def convert_videos(filename):
    with open(filename,'rb') as f:
        t = f.read().decode('utf-8-sig')
    my_yaml = yaml.load(t,Loader=yaml.FullLoader)
        
    root_directory = os.path.normpath(os.path.split(filename)[0])
#    video_directory = os.path.join(root_directory,'videos')
#    thumbs_directory = os.path.join(root_directory,'thumbs')
        
    #movie = {'file':'asdf.mov','asdf':'y'}
    #
    #slides = [[movie.copy(),movie.copy()],[movie.copy(),movie.copy()],None]
    #
    #with open('video_info_template.yaml','w') as f:
    #    yaml.dump(slides,f)
    
#    else:
#        shutil.rmtree(video_directory)
#        time.sleep(1)
#        os.mkdir(video_directory)
#
#
#    else:
#        shutil.rmtree(thumbs_directory)
#        time.sleep(1)
#        os.mkdir(thumbs_directory)
    
    if my_yaml is not None:
        for movie in my_yaml:
            path = movie['source']
            fn2 = movie['video_dest_path']
            image_name = movie['thumb_dest_path']
    
            video_filename = os.path.normpath(os.path.join(root_directory,fn2))
            video_directory = os.path.split(video_filename)[0]
            image_filename = os.path.normpath(os.path.join(root_directory,image_name))
            thumbs_directory = os.path.split(image_filename)[0]

            if not os.path.exists(video_directory):
                os.mkdir(video_directory)
            if not os.path.exists(thumbs_directory):
                os.mkdir(thumbs_directory)
    
            if not os.path.exists(video_filename):
                options = ''
                
                info = vi.VideoInfo(video_filename)

                if 'start_point' in movie:
                    tt1 = int(movie['start_point'])
                    t1 = '{0:.3f}'.format(tt1/1000)
                    options += '-ss '+t1+' '
                else:
                    t1=0
                
                
                if 'length' in movie:
                    tt4 = int(movie['length'])
                else:
                    tt4 = info.get_videos()[0].duration.float

 
                if 'end_point'  in movie:
                    tt2 = int(movie['end_point'])
                    
                    if abs((tt2-tt4)/tt4)>.01:
                        t2 = '{0:.3f}'.format(tt2/1000)
                        options += '-to '+t2+' '
               
    
    #                    s='ffmpeg -i "'+path+'" -ss '+t1+' -to '+t2+' -c:v libx264 -async 1 -preset ultrafast -crf 40 "'+video_filename+'"'                    
                # s='ffmpeg -i "'+path+'" ' + options + ' -c:v libx264 -async 1 -preset ultrafast -crf 30 "'+video_filename+'"'                    
                s='ffmpeg -i "'+path+'" ' + options + ' -c:v libx264 -async 1 -preset veryslow -crf 35 "'+video_filename+'"'                    
    #                    s='ffmpeg -i "'+path+'" -ss '+t1+' -to '+t2+' -c:v libx264 -async 1 -preset veryslow -crf 40 "'+video_filename+'"'                    
                print(s)
                b=subprocess.run(s, shell=True, capture_output=True)

                t3 = '{0:.3f}'.format((tt1+tt2)/2/1000)
                s3 = 'ffmpeg -ss '+t3+' -i "'+path+'" -frames:v 1 "'+image_filename+'"'
    #                    print(s3)

                b=subprocess.run(s3, shell=True, capture_output=True)
            else:
                # raise Exception('already exists: '+video_filename)
                print('already exists: '+video_filename)
                
if __name__=='__main__':
    
# #    directory = 'C:/Users/danaukes/projects/class_foldable_robotics/modules'
# #    directory = 'C:/Users/danaukes/projects/class_foldable_robotics/modules/01-introduction'
# #    directory = 'C:/Users/danaukes/projects/class_foldable_robotics/modules/02-bio-inspired robotics I - biomechanics & locomotion'
# #    directory = 'C:/Users/danaukes/projects/class_foldable_robotics/modules/05-foldable robotics background'
# #    directory = 'C:/Users/danaukes/projects/class_foldable_robotics/modules/07-kinematics I'
# #    directory = 'C:\\Users\\danaukes\\projects\\class_foldable_robotics\\modules\\upcoming\\19-bio-inspired-robots II - terrestrial locomotion'
# #    directory = 'C:\\Users\\danaukes\\projects\\class_foldable_robotics\\modules\\upcoming\\Rapid Prototyping & Laser Cutting'
# #    directory = 'C:\\Users\\danaukes\\Dropbox (ASU)\\idealab\\presentations\\2020-01 Research Talk'
# #    directory = 'C:/Users/danaukes/Desktop/testpres'
#     directory = 'C:/Users/danaukes/Dropbox (ASU)/idealab/presentations/2020-03-05 Research Talk/reduced'
# #    for dirpath,dirnames,filenames in os.walk(directory):
#     a=glob.glob(directory+'/**/*-video-info.yaml',recursive=True)

    a=[]
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\01-introduction\\lecture-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\02-bio-inspired-robotics-I-biomechanics-and-locomotion\\lecture-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\05-foldable-robotics-background\\lecture-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\07-kinematics-I\\lecture-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\19-mechatronics-I-arduino-and-servos\\lecture-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\24-dash-robotics\\lecture-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_drafts\\30-final-lecture\\lecture(2019)-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_upcoming\\19-bio-inspired-robots II - terrestrial locomotion\\13-bio-inspiration-video-info.yaml')
    # a.append('C:/Users/danaukes/projects/project_foldable_robotics/lectures\\_upcoming\\Rapid Prototyping & Laser Cutting\\rapid-prototyping-and-laser-cutting-video-info.yaml')
    a.append(r'C:\Users\danaukes\Dropbox (ASU)\idealab\presentations\2021-02-09 Seminar\seminar.yaml')

    for item in a:
        convert_videos(item)