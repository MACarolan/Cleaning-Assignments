# -*- coding: utf-8 -*-
"""
Created on Fri Oct 19 18:59:27 2018

@author: Michael
"""
from openpyxl import Workbook as wb
from openpyxl import load_workbook
import random
import sys

#name, year
bros = [('Deely', 19),
        ('Carter', 19),
        ('Murphy', 19),
        ('Mutty', 19),
        ('CJ', 15), #negative weight
        ('Popik', 19),
        ('Reilly', 19),
        ('Tim', 19),
        ('Seo', 19),
        ('Noah', 14),
        ('Dan', 16), #negative weight
        ('Vince', 19),
        ('Alden', 20),
        ('Dylan', 20),
        ('Kevin', 20),
        ('Tony', 20),
        ('Nic', 20),
        ('Hakeem', 20),
        ('Shubert', 20),
        ('Aidan', 20),
        ('Dillon', 21),
        ('Pablo', 21),
        ('Bowman', 21),
        ('Adam', 21),
        ('Nash', 21),
        ('Timi', 21),
        ('Paris', 21),
        ('Rostam', 21),
        ('Collin', 21),
        ('Jason', 21),
        ('Jonah', 22),
        ('Gio', 22),
        ('Cooper', 22),
        ('Liam', 22),
        ('Fiyi', 22),
        ('Santoro', 22),
        ('Ian', 22),
        ('Brian', 22),
        ('Kyri', 22)]

#assignment, difficulty 1-5
assignments = [('Laundry Room', 1),
               ('Pong Room', 3),
               ('Stalag', 3),
               ('Basement BR', 3),
               ('Bike Room', 3),
               ('Kitchen', 2),
               ('Freshman study', 2),
               ('Maidenhead', 2),
               ('Foyer', 2),
               ('Mail', 1),
               ('Library', 1),
               ('Pool Room', 2),
               ('Main Hall', 4),
               ('Dining Hall', 5),
               ('Pantry', 4),
               ('Gold Room', 2),
               ('Music Room', 2),
               ('3rd Landing', 4),
               ('3rd Bathroom', 3),
               ('3rd Shower', 2),
               ('4th Landing', 3),
               ('4th Shower', 2),
               ('4th Bathroom', 3),
               ('5th Hallway', 3),
               ('5th Bathroom', 2),
               ('Main Stairs 1-2', 1),
               ('Main Stairs 2-3', 1),
               ('Main Stairs 3-4', 1),
               ('Back Stairs 0-3', 1),
               ('Back Stairs 3-5', 1)]

doubles = ['Dining Hall', 'Pantry', '3rd Landing', '4th Landing']

def get_work(brother):
    """
    Returns cumulative work for each brother as int
    """
    log = open('LOG.txt', 'r').read()
    start, end = 0, 0
    for pos in range(len(log)-len(brother)+1): #find the brother's name as string
        if log[pos:pos+len(brother)] == brother and log[pos+len(brother)] == ':':
            start = pos+len(brother)+1
            end = start
            while log[end] != ',':
                end += 1
                
    if start == end:
        return None
    
    return int(log[start:end])

def set_work(brother, add):
    """
    Updates cumulative work for each brother
    """
    log = open('LOG.txt', 'r')
    lines = log.readlines()
    log = open('LOG.txt', 'w')
    
    for line in lines:
        
        if line[0:len(brother)] == brother and line[len(brother)] == ':':
            start = len(brother)+1
            end = start
            while line[end] != ',':
                end += 1
                
            line = brother + ':' + str(int(line[start:end]) + add) + ',\n'
        log.write(line)
    
    log.close()
    
def make_bro_list():
    '''
    Make list of bros ordered by work and then class year
    '''
    log = open('LOG.txt', 'r')
    for bro in range(len(bros)):
        bros[bro] = (bros[bro][0], bros[bro][1], get_work(bros[bro][0]))
    log.close()
    random.shuffle(bros)
    bros.sort(key = lambda b: b[2]-b[1])
    
def get_diff(job):
    for jobs in assignments:
        if job == jobs[0]:
            return jobs[1]
        
def swap_senior(pos):
    '''
    Switch a senior with a later freshman
    '''
    offset = 0
    while pos+offset < len(bros) and bros[pos+offset][1] <= 19:
        offset += 1
    if pos < len(bros) and pos+offset < len(bros):
        bros[pos], bros[pos+offset] = bros[pos+offset], bros[pos]
        
    
##potential updates, make floor buckets to pair residents w/ assignments
    
#load sheet
wb = load_workbook('Template.xlsx')
new = wb.active

##debugging, reset
reset = 0
if len(sys.argv) == 2:
    reset = int(sys.argv[1])
if reset:
    log = open('LOG.txt','w+')
    for bro in range(len(bros)):
        log.write(f'{bros[bro][0]}'+':0,\n')
    log.close()

#pick brothers with min work total by class year
if not reset:
    make_bro_list()
    count = 0
    
    places = list(range(2,len(assignments)+2))
    random.shuffle(places)
    for num in places:
        job = new[f'A{num}'].value
        diff = get_diff(job)
        cell = f'B{num}'
        if diff >= 4:
            swap_senior(count)
        brother = bros[count][0]
        
        if job not in doubles:
            new[cell] = brother
            set_work(brother, diff)
            count += 1
        else:
            if diff >= 4:
                swap_senior(count+1)
            new[cell] = brother + ', ' + bros[count+1][0]
            set_work(brother, diff)
            set_work(bros[count+1][0], diff)
            count += 2
    
    wb.save('assignments.xlsx')