# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 05:49:16 2020

@author: danaukes
"""

import glob
import os
import yaml


folder = 'C:/Users/danaukes/projects/project_foldable_robotics/lectures'
search_string = '/**/*-video-info.yaml'
full_path = folder+search_string
a=glob.glob(full_path,recursive=True)
for item in a:
    with open(item) as f:
        structure = yaml.load(f, Loader=yaml.FullLoader)
        for video in structure:
            filename = video['source']
            if not os.path.exists(filename):
                print(item,filename)

print(a)