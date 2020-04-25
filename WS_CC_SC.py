# -*- coding: utf-8 -*-
"""
Spyder Editor

Dies ist eine temporäre Skriptdatei.
"""
import os
import numpy as np
import sys
import math
import WS_CC_SC_Function as Func

np.set_printoptions(threshold=sys.maxsize)

### ---------------------------------- GET THE CURRENT WORKING DIRECTORY AND THE FILE NAME OF THE FILE ------------------------------------------------------- ###
filePath = os.getcwd()
excelName, loadCase, weldingType, CCEndBeam, CCMiddleBeam, SCEndBeam, SCMiddleBeam, REGEndBeam, REGMiddleBeam, comType = Func.ReadInputFileFromTextFile(filePath)
### -------------------------------------------------------------------------------------------------------------------------------------------- ###
#print(loadCase[1])
### ----------------------------------- INITIALIZE VARIABLE TO STORE THE VALUE WHILE READING THE FILE ------------------------------------------------------
for Loop in range(len(loadCase)):
    
#    print(loadCase[Loop])
    elementID = []
    elementCentroid = []
    componentID = []
    componentName = []
    contourValue = []

### ------------------------------------------------------------ READ THE FILE ------------------------------------------------------------------------------
    elementID, elementCentroid, componentID, componentName, contourValue = Func.ReadInputFileFromHyperworks(loadCase[Loop])
### ------------------------------------------------------------ READ THE FILE ------------------------------------------------------------------------------
    print(len(elementID))
### ---------------------------------------------------------- COMPUTE AVERAGE ELEMENT CENTROID -------------------------------------------------------------

    xCoordinate, yCoordinate, zCoordinate = Func.SplitElementCentroid(elementCentroid)

    sortedxCoordinate = np.zeros((len(xCoordinate)))
    sortedyCoordinate = np.zeros((len(xCoordinate)))
    sortedzCoordinate = np.zeros((len(xCoordinate)))
    sortedComponentName = ["" for x in range (len(componentName))]
    sortedContourValue = np.zeros((len(contourValue)))

### ---------------------------------------------------------- COMPUTE AVERAGE ELEMENT CENTROID -------------------------------------------------------------
    if Loop == 0:
### --------------------------------------------------- SORTING THE VARIABLE VALUE --------------------------------------------------------------------------

        
        countComponentName = (sorted(set(componentName))) #####Count the number of component (usually 3: LH, MID, RH)#######

        sortedElementID = sorted(elementID)

        indexSortedElementID = sorted (range(len(elementID)), key=elementID.__getitem__) ##Index of sorted element ID##


  
        for i in range (0, len(sortedElementID)):
    
            sortedxCoordinate[i] = xCoordinate[indexSortedElementID[i]]
            sortedyCoordinate[i] = yCoordinate[indexSortedElementID[i]]
            sortedzCoordinate[i] = zCoordinate[indexSortedElementID[i]]
            sortedContourValue[i] = contourValue[indexSortedElementID[i]]
            sortedComponentName[i] = componentName[indexSortedElementID[i]]

        Func.WriteSortedCoordinate(filePath, sortedxCoordinate, sortedyCoordinate, sortedzCoordinate)
        Func.WriteSortedVariableIntoTextFile(sortedElementID, sortedxCoordinate, sortedyCoordinate, sortedzCoordinate, sortedContourValue, sortedComponentName)
    
    else:
        index, sortedxCoordinate, sortedyCoordinate, sortedzCoordinate = Func.MatchElementCentroid(filePath+'\Coordinate.txt', xCoordinate, yCoordinate, zCoordinate)
        print(len(index))
        sortedIndex = sorted(index)
        print(sortedIndex)
        sortedElementID = np.zeros((len(sortedxCoordinate)))
        with open('loopCoordinate.txt','w') as f:
            
            for i in range (len(sortedxCoordinate)):
                f.write(str(sortedxCoordinate[i])+' ')
                f.write(str(sortedyCoordinate[i])+' ')
                f.write(str(sortedzCoordinate[i])+'\n')
        
        for i in range (len(sortedxCoordinate)):
            sortedElementID[i]= elementID[index[i]]
            sortedContourValue[i] = contourValue[index[i]]
            sortedComponentName[i] = componentName[index[i]]
#### ------------------------- Initialize the array needed for the classification of the weld seam locations ---------------------------------------------
#        
    elementIDLeft = []
    elementIDMiddle = []
    elementIDRight = []
    elementIDLocation = [elementIDLeft, elementIDMiddle, elementIDRight]

    contourValueLeft = []
    contourValueMiddle = []
    contourValueRight = []
    contourValueLocation = [contourValueLeft, contourValueMiddle, contourValueRight]

    xCentroidLeft = []
    xCentroidMiddle = []
    xCentroidRight = []  
    xCentroidLocation = [xCentroidLeft, xCentroidMiddle, xCentroidRight]
    
    yCentroidLeft = []
    yCentroidMiddle = []
    yCentroidRight = []  
    yCentroidLocation = [yCentroidLeft, yCentroidMiddle, yCentroidRight]
    
    zCentroidLeft = []
    zCentroidMiddle = []
    zCentroidRight = []  
    zCentroidLocation = [zCentroidLeft, zCentroidMiddle, zCentroidRight]
#
    valueDifferenceXLeft = []
    valueDifferenceXMiddle = []
    valueDifferenceXRight = []
    valueDifferenceXLocation = [valueDifferenceXLeft, valueDifferenceXMiddle, valueDifferenceXRight]

    valueDifferenceYLeft = []
    valueDifferenceYMiddle = []
    valueDifferenceYRight = []
    valueDifferenceYLocation = [valueDifferenceYLeft, valueDifferenceYMiddle, valueDifferenceYRight]

    valueDifferenceZLeft = []
    valueDifferenceZMiddle = []
    valueDifferenceZRight = []
    valueDifferenceZLocation = [valueDifferenceZLeft, valueDifferenceZMiddle, valueDifferenceZRight]

    valuexDifference = np.zeros((len(sortedElementID)))
    valueyDifference = np.zeros((len(sortedElementID)))
    valuezDifference = np.zeros((len(sortedElementID)))
    #
### ------------------ Store the value of elementID, elementCentroid, and contour value into their locations (left, middle, or right) --------------------
    for i in range (0,len(countComponentName)):
    
        for j in range (0,len(elementID)):
        
            if sortedComponentName[j] == countComponentName[i]:
            
                elementIDLocation[i].append(sortedElementID[j])
                xCentroidLocation[i].append(sortedxCoordinate[j])
                yCentroidLocation[i].append(sortedyCoordinate[j])
                zCentroidLocation[i].append(sortedzCoordinate[j])
                contourValueLocation[i].append(sortedContourValue[j])

    for i in range (len(valueDifferenceXLocation)):

        valueDifferenceXLocation[i] = np.around(abs(np.diff(xCentroidLocation[i]).astype(float)),decimals=1)
        valueDifferenceYLocation[i] = np.around(abs(np.diff(yCentroidLocation[i]).astype(float)),decimals=1)
        valueDifferenceZLocation[i] = np.around(abs(np.diff(zCentroidLocation[i]).astype(float)),decimals=1)
   
        with open('valDiffX.txt','w') as f:
#    f.write(str(xDifference)+' ')
            for i in range (len(valueDifferenceXLocation[0])):
                f.write(str(valueDifferenceXLocation[0][i])+' ')
                f.write(str(valueDifferenceYLocation[0][i])+' ')
                f.write(str(valueDifferenceZLocation[0][i])+'\n')
#### ----------------------------------------------------------------------------------------------------------------------------------------------------
#In here, class of computation of number of element in a weld seam in left,middle,and right hand side can be applied
    for i in range (len(countComponentName)):
    

#for i in range(len(valueDifferenceLocation)):
        ElList=[]
        ContourList=[]
        valueEndBeam=[]
        valueMiddleBeam=[]
        resultStatus=[]
        ElList, ContourList = Func.DetermineListofWeldSeaminOneStructure(elementIDLocation[i], contourValueLocation[i], valueDifferenceXLocation[i], valueDifferenceYLocation[i], valueDifferenceZLocation[i])

####### Start the computation of CC SC Definition #########################################
        
#        weldSeamType ='Laser'

        valueEndBeam, valueMiddleBeam, resultStatus = Func.CCandSCDefinition(weldingType, CCEndBeam, CCMiddleBeam, SCEndBeam, SCMiddleBeam, REGEndBeam, REGMiddleBeam, ContourList, comType)
#        print(valueMiddleBeam)
        variable=['WeldSeamElement','ValueEndBeam','ValueMiddleBeam','Status']
        value=[ElList, valueEndBeam, valueMiddleBeam, resultStatus]

############################################################ Write the result into excel file ###########################################################
        if Loop == 0:
            if i != len(countComponentName)-1:
                if i == 0:
        
                    Func.SaveResultToExcel (excelName, loadCase[Loop], 1, 2, countComponentName[i], variable, value, CCEndBeam, CCMiddleBeam, SCEndBeam, SCMiddleBeam, REGEndBeam, REGMiddleBeam)
                    with open('ComponentName.txt','w') as f:
                        f.write(str(countComponentName[i])+',')
            
                else:
        
                    Func.RewriteExcelWithoutSavingRow (excelName, loadCase[Loop], 1, 2, countComponentName[i], variable, value, CCEndBeam, CCMiddleBeam, SCEndBeam, SCMiddleBeam, REGEndBeam, REGMiddleBeam)
                    with open('ComponentName.txt','a') as f:
                        f.write(str(countComponentName[i])+',')
            else:
         
                Func.RewriteExcel(excelName, loadCase[Loop], 1, 2, countComponentName[i], variable, value, CCEndBeam, CCMiddleBeam, SCEndBeam, SCMiddleBeam, REGEndBeam, REGMiddleBeam)
                with open('ComponentName.txt','a') as f:
                    f.write(str(countComponentName[i]))
 
        else:

            if i != len(countComponentName)-1:
                lastRow, firstColumn= Func.ReadLastRow(filePath)
                latestRow=int(lastRow)
    
                Func.LoopWithoutSavingLatestRow(excelName, loadCase[Loop], 1, latestRow+1, countComponentName[i], variable, value)
    
            else:
                lastRow, firstColumn= Func.ReadLastRow(filePath)
                latestRow=int(lastRow)
                
                Func.LoopWithSavingLatestRow(excelName, loadCase[Loop], 1, latestRow+1, countComponentName[i], variable, value)

with open ('ComponentName.txt','r') as r:
    
    compoName = []
    lines= r.readlines()
    
    for line in lines:
        linelist = line.split(',')
        for i in range (len(linelist)):
           
            compoName.append(linelist[i])

with open ('ColumnLength.txt','r') as r:
    
    columnLength =[]
    lines= r.readlines()
    
    for line in lines:
        linelist = line.split()
        columnLength.append(linelist[1])

with open ('RowContainingResult.txt','r') as r:
    
    resultRow =[]
    lines= r.readlines()
    
    for line in lines:
        linelist = line.split()
        resultRow.append(linelist[1])
        
last, first= Func.ReadLastRow(filePath)

for i in range (len(compoName)):
    Func.ComputeFinalStatusExcel(excelName, compoName[i], int(columnLength[i]), int(last)+2, int(first)+1, resultRow)
###if it issecond loop then do:
        
        ###readfrom the text file to fetch the value of the last row, then use it as an input argument for rewrite excel