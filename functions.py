import numpy as np
import matplotlib.pyplot as plt
import math
import os


def rotate_points(point_x, point_y, center_x, center_y, angle_deg, scalingFactor = False):
    angle_rad = math.radians(angle_deg)  # Convert angle to radians


    if scalingFactor:
        final_rotated_px = center_x + scalingFactor * (math.cos(angle_rad) * (point_x - center_x) - math.sin(angle_rad) * (point_y - center_y))
        final_rotated_py = center_y + scalingFactor * (math.sin(angle_rad) * (point_x - center_x) + math.cos(angle_rad) * (point_y - center_y))

    else:
        final_rotated_px = center_x + math.cos(angle_rad) * (point_x - center_x) - math.sin(angle_rad) * (point_y - center_y)
        final_rotated_py = center_y + math.sin(angle_rad) * (point_x - center_x) + math.cos(angle_rad) * (point_y - center_y)
    return final_rotated_px, final_rotated_py



def getOctagonFractalUnit(numberOfUnits, biggestRadius, scalingFactor, center_x, center_y):
    globalList = list()
    # Circle parameters
    shrinkFromSecondToThird = False
    radius = biggestRadius
    rotationDegree = 45 / 2
    for numberOfUnit in range(numberOfUnits):
        if numberOfUnit == 0: #this means that there are only one component in the fractal
            
            xypairsOuterCircle = list()
            xypairsInnerCircle = list()

            # Calculate points
            angle_deg = 45 * 0
            angle_rad = math.radians(angle_deg)  # Convert to radians
            firstpoint_x = center_x + radius * math.cos(angle_rad)
            firstpoint_y = center_y + radius * math.sin(angle_rad)
            listToAppend = [center_x, center_y, firstpoint_x, firstpoint_y]
            
            points_x, points_y = [], []
            for i in range(1, 8):
                angle_deg = 45 * i
                angle_rad = math.radians(angle_deg)  # Convert to radians
                point_x = center_x + radius * math.cos(angle_rad)
                point_y = center_y + radius * math.sin(angle_rad)
                listToAppend.append(point_x)
                listToAppend.append(point_y)
                xypairsOuterCircle.append(listToAppend)
                # print(listToAppend) 
                
                rotated_points = [(center_x, center_y) if ((listToAppend[pointIndex] == center_x) and (listToAppend[pointIndex + 1] == center_y)) else rotate_points(listToAppend[pointIndex], listToAppend[pointIndex + 1], center_x, center_y, rotationDegree, scalingFactor) for pointIndex in range(0, len(listToAppend), 2)]
                flat_list = [value for pair in rotated_points for value in pair]
                xypairsInnerCircle.append(flat_list)

                listToAppend = [center_x, center_y, point_x, point_y]
                # print(rotated_points)

            listToAppend.append(firstpoint_x)
            listToAppend.append(firstpoint_y)
            # print(listToAppend) 
            xypairsOuterCircle.append(listToAppend)

            rotated_points = [(center_x, center_y) if ((listToAppend[pointIndex] == center_x) and (listToAppend[pointIndex + 1] == center_y)) else rotate_points(listToAppend[pointIndex], listToAppend[pointIndex + 1], center_x, center_y, rotationDegree, scalingFactor) for pointIndex in range(0, len(listToAppend), 2)]
            # print(rotated_points) 
                
            flat_list = [value for pair in rotated_points for value in pair]
            xypairsInnerCircle.append(flat_list)

            globalList.append(xypairsOuterCircle)
            # print(len(xypairsOuterCircle))
            globalList.append(xypairsInnerCircle)
            # print(len(xypairsInnerCircle))

        else:
            xypairsOuterCircle = list()
            if shrinkFromSecondToThird:
                for points in xypairsInnerCircle:
                    listToAppend = list()
                    # print('points', len(points))
                    for xypairindex in range(0, len(points), 2):
                        
                        rotated_px, rotated_py = rotate_points(points[xypairindex] * scalingFactor, points[xypairindex + 1] * scalingFactor, center_x, center_y, rotationDegree)
                        listToAppend.append(rotated_px)
                        listToAppend.append(rotated_py)
                    xypairsOuterCircle.append(listToAppend)
            else:
                for points in xypairsInnerCircle:
                    
                    listToAppend = list()
                    for xypairindex in range(0, len(points), 2):
                        if ((points[xypairindex] == center_x) and (points[xypairindex + 1] == center_y)):
                            # print('not rotated')
                            rotated_px, rotated_py = center_x, center_y
                        else:
                            rotated_px, rotated_py = rotate_points(points[xypairindex], points[xypairindex + 1], center_x, center_y, rotationDegree)
                        listToAppend.append(rotated_px)
                        listToAppend.append(rotated_py)
                    xypairsOuterCircle.append(listToAppend)
            
            xypairsInnerCircle = list()
            for points in xypairsOuterCircle:
                    listToAppend = list()
                    for xypairindex in range(0, len(points), 2):
                        if ((points[xypairindex] == center_x) and (points[xypairindex + 1] == center_y)):
                            rotated_px, rotated_py = center_x, center_y
                        else:
                            rotated_px, rotated_py = rotate_points(points[xypairindex], points[xypairindex + 1], center_x, center_y, rotationDegree, scalingFactor)
                        listToAppend.append(rotated_px)
                        listToAppend.append(rotated_py)
                    xypairsInnerCircle.append(listToAppend)

            globalList.append(xypairsOuterCircle)
            # print(len(xypairsOuterCircle))
            globalList.append(xypairsInnerCircle)
            # print(len(xypairsInnerCircle))

    return globalList, center_x, center_y


# Function to draw triangles
def draw_triangles(globalList):
    
    plt.figure(figsize=(4, 4))
    colors = ['green', 'orange', 'cyan', 'purple', 'black', 'red']
    for trianglesIndex, triangles in enumerate(globalList):
        for triangle in triangles:
            # Extract x and y pairs
            x_coords = [triangle[i] for i in range(0, len(triangle), 2)]
            y_coords = [triangle[i + 1] for i in range(0, len(triangle), 2)]
            
            # Close the triangle by adding the first point again at the end
            x_coords.append(x_coords[0])
            y_coords.append(y_coords[0])
            
            # Plot the triangle

            plt.plot(x_coords, y_coords, marker='o',color = colors[trianglesIndex])  # Line and points
        # color = 'red'
    plt.axhline(0, color='black', linewidth=0.5, linestyle='--')
    plt.axvline(0, color='black', linewidth=0.5, linestyle='--')
    plt.grid(True)
    plt.title("Triangles from Coordinate Data")
    plt.xlabel("X-axis")
    plt.ylabel("Y-axis")
    plt.axis('equal')  # Ensure equal scaling for both axes
    

def line3(pointNum1, pointNum2, step, antennaMaterial):
     lines_to_insert3 = [f'For m={pointNum1} To {pointNum2} STEP {step}\n',
     'With Polygon\n',
          '     .Reset \n',
          '     .Name "polygon1" +Str(m) \n',
          '     .Curve "curve1" \n',
          '     .Point x(m), y(m) \n',
          '     .LineTo x(m+1), y(m+1) \n',
          '     .LineTo x(m+2), y(m+2) \n',
          '     .LineTo x(m), y(m) \n',
          '     .Create \n',
     'End With \n',
     
     "With ExtrudeCurve\n",
     "     .Reset \n",
     '     .Name "solid" + Str(m) \n',
     '     .Component "component1" \n',
     f'     .Material "{str(antennaMaterial)}" \n',
     f'     .Thickness "{thickness}" \n',
     '     .Twistangle "0.0"\n' ,
     '     .Taperangle "0.0"\n',
     '     .DeleteProfile "True"\n' ,
     '     .Curve "curve1:polygon1" + Str(m)\n' ,
     '     .Create\n',
     'End With\n',
     'Next m \n'
     ]
     lines_to_insert3 = ''.join(lines_to_insert3)
     return lines_to_insert3


def line4(added, subsfrom1):
    lines_to_insert4 = ["'## Merged Block - boolean subtract shapes: component1:patch, component1: remove\n",
    'StartVersionStringOverrideMode "2024.4|33.0.1|20240430"\n', 
    f'Solid.Subtract "component1:solid {subsfrom1}", "component1:solid {added}"\n']
    lines_to_insert4 = ''.join(lines_to_insert4)
    return lines_to_insert4

def line5(added, subsfrom2):
    lines_to_insert5 = ["'## Merged Block - boolean subtract shapes: component1:patch, component1: remove\n",
    'StartVersionStringOverrideMode "2024.4|33.0.1|20240430"\n', 
    f'Solid.Subtract "component1:solid {subsfrom2}", "component1:solid {added}"\n']

    lines_to_insert5 = ''.join(lines_to_insert5)
    return lines_to_insert5


def engine(Start_Frequency, End_Frequency, Thickness, biggest_circle_radius, Number_of_layers, Scaling_Ratio, SubstrateLength, SubstrateWidth, groundThickness, SubstrateHeight, Dipole_Connection_Length, Dipole_Connection_Width, Feedline_length, Feedline_width, sToBeDefined, antennaMaterial, substrateMaterial, groundMaterial, Input_FileName_to_modify, Output_FileName_to_Save):

    for pole in range(2):

        global thickness
        global startFrequency
        global endFrequency
        width_, length_ = 6*SubstrateHeight + biggest_circle_radius*2 + Dipole_Connection_Length, 6*SubstrateHeight + biggest_circle_radius*4

        original_file = f'{Input_FileName_to_modify}.mcs'  # Replace with the actual file name
        # new_file = f'{Output_FileName_to_Save}.mcs'  # New file name
        new_file = f'{Output_FileName_to_Save}.mcs'

        startFrequency = Start_Frequency
        endFrequency = End_Frequency
        thickness = Thickness

        # draw_triangles(globalList)


        lines_to_insert1 = [
            "'## Merged Block - define curve polygon: curve1:polygon1\n",
            'StartVersionStringOverrideMode "2024.4|33.0.1|20240430"\n'
        ]
        lines_to_insert1 = ''.join(lines_to_insert1)

        midPointx, midPointy = 0, 0 

        h = (Dipole_Connection_Width / 2) / math.tan(math.radians(67.5))
        center_x1, center_y1 = midPointx - (Dipole_Connection_Length / 2) + (h) - biggest_circle_radius, midPointy  # Center coordinates
        globalList1, center_x1, center_y1 = getOctagonFractalUnit(Number_of_layers, biggest_circle_radius, Scaling_Ratio, center_x1, center_y1)
        center_x2, center_y2 = midPointx + (Dipole_Connection_Length / 2)  - (h) + biggest_circle_radius, midPointy # Center coordinates
        globalList2, center_x2, center_y2 = getOctagonFractalUnit(Number_of_layers, biggest_circle_radius, Scaling_Ratio, center_x2, center_y2)
        
        lines_to_insert2 = list()
        pointNum = len(globalList1)*len(globalList1[0]) * 3 + len(globalList2)*len(globalList2[0]) * 3
        lines_to_insert2.append(f'Dim x({pointNum + 4 + 4}), y({pointNum + 4+ 4}) As Double\n')
        lines_to_insert2.append(f'Dim m As Integer\n')
        
        center_x = (center_x1 + center_x2) / 2
        center_y = (center_y1 + center_y2) / 2

        pointCounter = 1
        for globalList in [globalList1, globalList2]:
            
            for octagon in globalList:
                for threePointList in octagon:
                    for xyindex in range(0, len(threePointList), 2):
                        x = threePointList[xyindex]
                        y = threePointList[xyindex + 1]
                        lines_to_insert2.append(f'x({pointCounter})={x}\n')
                        lines_to_insert2.append(f'y({pointCounter})={y}\n')
                        pointCounter += 1

        #######this is for interpoles connection line
        # connectionx1, connectiony1 = midPointx - (Dipole_Connection_Length / 2) - ((Dipole_Connection_Width / 2) / math.tan(math.radians(67.5))), midPointy + Dipole_Connection_Width / 2
        connectionx1, connectiony1 = midPointx - (Dipole_Connection_Length / 2), midPointy + Dipole_Connection_Width / 2
        connectionFirstCounter = pointCounter
        lines_to_insert2.append(f'x({pointCounter})={connectionx1}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony1}\n')
        pointCounter += 1

        # connectionx2, connectiony2 = midPointx - (Dipole_Connection_Length / 2) - ((Dipole_Connection_Width / 2) / math.tan(math.radians(67.5))), midPointy - Dipole_Connection_Width / 2
        connectionx2, connectiony2 = midPointx - (Dipole_Connection_Length / 2), midPointy - Dipole_Connection_Width / 2
        lines_to_insert2.append(f'x({pointCounter})={connectionx2}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony2}\n')
        pointCounter += 1

        # connectionx3, connectiony3 = midPointx + (Dipole_Connection_Length / 2) + ((Dipole_Connection_Width / 2) / math.tan(math.radians(67.5))), midPointy - Dipole_Connection_Width / 2
        connectionx3, connectiony3 = midPointx + (Dipole_Connection_Length / 2), midPointy - Dipole_Connection_Width / 2
        lines_to_insert2.append(f'x({pointCounter})={connectionx3}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony3}\n')
        pointCounter += 1

        # connectionx4, connectiony4 = midPointx + (Dipole_Connection_Length / 2) + ((Dipole_Connection_Width / 2) / math.tan(math.radians(67.5))), midPointy + Dipole_Connection_Width / 2
        connectionx4, connectiony4 = midPointx + (Dipole_Connection_Length / 2), midPointy + Dipole_Connection_Width / 2
        lines_to_insert2.append(f'x({pointCounter})={connectionx4}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony4}\n')
        pointCounter += 1


        #######this is for feedline
        connectionx1, connectiony1 = midPointx - (Feedline_width / 2), midPointy
        connectionFirstCounter2 = pointCounter
        lines_to_insert2.append(f'x({pointCounter})={connectionx1}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony1}\n')
        pointCounter += 1

        connectionx2, connectiony2 = midPointx - (Feedline_width / 2), center_y - (width_ / 2)#Feedline_length
        lines_to_insert2.append(f'x({pointCounter})={connectionx2}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony2}\n')
        pointCounter += 1

        connectionx3, connectiony3 = midPointx + (Feedline_width / 2), center_y - (width_ / 2)#Feedline_length
        lines_to_insert2.append(f'x({pointCounter})={connectionx3}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony3}\n')
        pointCounter += 1

        connectionx4, connectiony4 = midPointx + (Feedline_width / 2), midPointy
        lines_to_insert2.append(f'x({pointCounter})={connectionx4}\n')
        lines_to_insert2.append(f'y({pointCounter})={connectiony4}\n')
        pointCounter += 1



        lines_to_insert2 = ''.join(lines_to_insert2)
####### Until here the coordinates of the polygons are defined and the points are stored in x and y arrays. ######
        
        pointCounter = 1
        combinedAdd = '' 
        for globalListIndex, globalList in enumerate([globalList1, globalList2]):
            
            for octagonIndex, octagon in enumerate(globalList):
                if (globalListIndex == 1):
                    octagonIndex = octagonIndex + len(globalList1)
                
                if octagonIndex % 2 == 0:
                    lines_to_insert3 = line3(octagonIndex*24 + 1, (octagonIndex + 1)*24, step = 3, antennaMaterial = antennaMaterial)
                    combinedAdd = combinedAdd + lines_to_insert3
                else:
                    for threePointListIndex in range(len(octagon)):
                        lines_to_insert3 = line3((octagonIndex * 24 + threePointListIndex * 3 + 1), (octagonIndex * 24 + (threePointListIndex + 1) * 3), step = 3, antennaMaterial = antennaMaterial)
                        combinedAdd = combinedAdd + lines_to_insert3
                    
                        lines_to_insert4 = line4((octagonIndex * 24 + threePointListIndex * 3 + 1), (octagonIndex * 24 - 24 + threePointListIndex * 3 + 1))
                        if threePointListIndex == (len(octagon) - 1):
                            lines_to_insert5 = line5((octagonIndex * 24 + threePointListIndex * 3 + 1), (octagonIndex * 24 + threePointListIndex * 3 + 1 - 48 + 3))
                        else:
                            lines_to_insert5 = line5((octagonIndex * 24 + threePointListIndex * 3 + 1), (octagonIndex * 24 - 24 + (threePointListIndex + 1) * 3 + 1))
                        combinedAdd = combinedAdd + lines_to_insert4
                        combinedAdd = combinedAdd + lines_to_insert3
                        combinedAdd = combinedAdd + lines_to_insert5


        print('connectionFirstCounter: ', connectionFirstCounter)
        

        lines_to_insert3 = [f'For m={connectionFirstCounter} To {connectionFirstCounter2 + 2} STEP 4\n',
        'With Polygon\n',
            '     .Reset \n',
            '     .Name "polygon1" +Str(m) \n',
            '     .Curve "curve1" \n',
            '     .Point x(m), y(m) \n',
            '     .LineTo x(m+1), y(m+1) \n',
            '     .LineTo x(m+2), y(m+2) \n',
            '     .LineTo x(m+3), y(m+3) \n',
            '     .LineTo x(m), y(m) \n',
            '     .Create \n',
        'End With \n',
        
        "With ExtrudeCurve\n",
        "     .Reset \n",
        '     .Name "solid" + Str(m) \n',
        '     .Component "component1" \n',
        f'     .Material "{str(antennaMaterial)}" \n',
        f'     .Thickness "{thickness}" \n',
        '     .Twistangle "0.0"\n' ,
        '     .Taperangle "0.0"\n',
        '     .DeleteProfile "True"\n' ,
        '     .Curve "curve1:polygon1" + Str(m)\n' ,
        '     .Create\n',
        'End With\n',
        'Next m \n'
        ]
        lines_to_insert3 = ''.join(lines_to_insert3)
        combinedAdd = combinedAdd + lines_to_insert3


        lines_to_insert = lines_to_insert1 + lines_to_insert2 + combinedAdd
        pickFaceLines = ["'## Merged Block - pick face\n",
'StartVersionStringOverrideMode "2024.4|33.0.1|20240430"\n',
f'Pick.PickFaceFromId "component1:solid {connectionFirstCounter2}", "3"\n']
        pickFaceLinesJoined = ''.join(pickFaceLines)
        lines_to_insert = lines_to_insert + pickFaceLinesJoined

        # Define file paths

        try:
            # Step 1: Open and read the original file
            with open(original_file, 'r') as file:
                content = file.readlines()  # Read lines into a list

            # Step 2: Make changes to the content
            new_content = [] = []
            
            for lineIndex, line in enumerate(content):
                if 'startFrequency' in line:
                    line = line.replace('startFrequency', str(startFrequency))
                    line = line.replace('endFrequency', str(endFrequency))
                    print('replaced frequencies')
                if 'Xrange' in line:
                    #SubstrateLength, SubstrateWidth, groundThickness, SubstrateHeight
                    line = line.replace('SubstrateLengthLower', str(center_x - (length_ / 2)))
                    line = line.replace('SubstrateLengthUpper', str(center_x + (length_ / 2)))
                if 'Yrange' in line:
                    line = line.replace('SubstrateWidthLower', str(center_y - (width_ / 2)))
                    line = line.replace('SubstrateWidthUpper', str(center_y + (width_ / 2)))
                if '.Zrange "-groundThicknessLower"' in line:
                    line = line.replace('-groundThicknessLower', str(-(groundThickness + SubstrateHeight)))
                    line = line.replace('groundThicknessUpper', str(-(SubstrateHeight)))
                if '.Zrange "-SubstrateHeight"' in line:
                    line = line.replace('-SubstrateHeight', str(-SubstrateHeight))
                
                if 'sDefineAt =' in line:
                    line = line.replace('sToBeDefined', str(sToBeDefined))

                if 'sDefineAtName = ' in line:
                    line = line.replace('sToBeDefined', str(sToBeDefined))
                
                if '.Material "groundMaterial" ' in line:
                    line = line.replace('groundMaterial', groundMaterial)
                if '.Material "substrateMaterial" ' in line:
                    line = line.replace('substrateMaterial', substrateMaterial)


                new_content.append(line)

                catchPhrase = 'HERE STARTS THE ANTENNA MODELLING'
                if catchPhrase in line:
                    new_content.extend(lines_to_insert)
                

            # Step 3: Save the modified content to a new file
            with open(new_file, 'w') as file:
                file.writelines(new_content)

            print(f"File modified and saved as: {new_file}")

        except FileNotFoundError:
            print(f"File not found: {original_file}")
        
        except Exception as e:
            print(f"An error occurred: {e}")

def constructPort(k_coefficient, portShieldMaterial, Output_FileName_to_Save):
    k_coefficient = k_coefficient.strip()
    port_lines_to_insert = ["'## Merged Block - define port:1\n",
                        'StartVersionStringOverrideMode "2024.4|33.0.1|20240430"\n',
                        "'Port constructed by macro Solver -> Ports -> Calculate port extension coefficient\n",
                        'With Port\n',

                        '  .Reset\n',

                        '  .PortNumber "1"\n',

                        '  .NumberOfModes "1"\n',

                        '  .AdjustPolarization False\n',

                        '  .PolarizationAngle "0.0"\n',

                        '  .ReferencePlaneDistance "0"\n',

                        '  .TextSize "50"\n',

                        '  .Coordinates "Picks"\n',

                        '  .Orientation "Positive"\n',

                        '  .PortOnBound "True"\n',

                        '  .ClipPickedPortToBound "False"\n',

                        f'  .XrangeAdd "1.5*{k_coefficient}", "1.5*{k_coefficient}"\n',

                        '  .YrangeAdd "0", "0"\n',

                        f'  .ZrangeAdd "1.5", "1.5*{k_coefficient}"\n',

                        f'  .Shield "{portShieldMaterial}"\n',

                        '  .SingleEnded "False"\n',

                        '  .Create\n',
                        'End With\n']
    
    port_lines_to_insert = ''.join(port_lines_to_insert)



    try:

        # Step 1: Open and read the original file
        with open(f'{Output_FileName_to_Save}.mcs', 'r') as file:
            content = file.readlines()  # Read lines into a list

        # print('new_filePath: ', Output_FileName_to_Save)

        # Step 2: Make changes to the content
        new_content = [] = []
        for lineIndex, line in enumerate(content):
            catchPhrase = 'Pick.PickFaceFromId'
            new_content.append(line)

            if catchPhrase in line:
                new_content.extend(port_lines_to_insert)
        # Obtain the parent directory
        parent_directory = os.path.dirname(Output_FileName_to_Save)

        # Ensure the parent directory ends with a slash
        if not parent_directory.endswith('/'):
            parent_directory += '/'

        new_fileName = 'simulation.mcs'
        new_filePath = f'{parent_directory}{new_fileName}'  # New file name    
        
        # Step 3: Save the modified content to a new file
        with open(new_filePath, 'w') as file:
            file.writelines(new_content)

        print(f"Simulation file saved as: {new_fileName}")
        return True, new_filePath
    
    except FileNotFoundError:
        print(f"File not found: {Output_FileName_to_Save}")
    
    except Exception as e:
        print(f"An error occurred: {e}")

    