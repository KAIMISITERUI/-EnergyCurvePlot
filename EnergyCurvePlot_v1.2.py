

import numpy as np
import matplotlib.pyplot as plt
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.widgets import Slider
import tkinter.colorchooser as colorchooser
import json
from tkinter import messagebox
import os
import math
from tkinter import simpledialog, filedialog
import pandas as pd
import tksheet

'''
1.3 开发计划
1. 预期加入excel支持功能，而不是单纯的json存储
2. 加入针对配置文件的保存功能

1.2 Update log

1. Fixed the bug where circular patterns appeared below the line layer due to identical IDs.
2. Fixed the bug where adjusting the page size affected the display (you can now freely modify the page size and aspect ratio).
3. Optimized the page adjustment logic to make it more reasonable.
4. Re-adjusted the bottom toolbar code; now all tools are functioning properly.
5. When the base curve = 0, the drawn shape is now a true straight line.
6 Added a right-click menu to the table area; right-clicking the table allows you to delete rows/columns, add rows on the left side, and add columns at the top.
7. Added a "show grid" option, allowing you to set whether to display gridlines.
8. Fixed the issue where exporting data from a table with empty rows caused an error.
9. Added an "Open Table" button to open the table area.
10. If multiple markers are at the same position, only one will be drawn.

Note: The table_data.json saved in all versions is universal and can be read by any version.

1.1 Update log

1. Auto-extend Page: The page now automatically extends, allowing for seamless content addition.
2. Set Label Position: Users can now customize the position of labels.
3. Single Marker Drawing: A feature has been added to draw a single marker on the plot for visual emphasis.
4. Font Size Customization: Users can now adjust the font size for text elements on the plot.
5. Font and Label Positioning: Added functionality to set the distance between the font and the label's central position.
6. Line Thickness Customization: Users can now adjust the thickness of lines drawn on the plot.
7. Bond Thickness Customization:  he thickness of bonds in the structure can now be customized.

1.0 veresion:
EnergyCurvePlot is a tool designed to generate energy change curves for chemical reactions. 
The program allows users to visualize the energy evolution throughout the reaction process, 
providing insights into reaction mechanisms and energy profiles. The output is a cdxml file, 
which can be easily opened and edited in Chemdraw for further customization. 

Powered by Yatao Lang at Lanzhou university
'''



cdxml_header = '''<?xml version="1.0" encoding="UTF-8" ?>
<!DOCTYPE CDXML SYSTEM "http://www.cambridgesoft.com/xml/cdxml.dtd" >
<CDXML
 CreationProgram="ChemDraw 22.0.0.22"
 Name="output.cdxml"
 BoundingBox="259.57 429.79 816.37 719.79"
 WindowPosition="0 0"
 WindowSize="-2147483648 0"
 WindowIsZoomed="yes"
 FractionalWidths="yes"
 InterpretChemically="yes"
 ShowAtomQuery="yes"
 ShowAtomStereo="no"
 ShowAtomEnhancedStereo="yes"
 ShowAtomNumber="no"
 ShowResidueID="no"
 ShowBondQuery="yes"
 ShowBondRxn="yes"
 ShowBondStereo="no"
 ShowTerminalCarbonLabels="no"
 ShowNonTerminalCarbonLabels="no"
 HideImplicitHydrogens="no"
 LabelFont="3"
 LabelSize="10"
 LabelFace="96"
 CaptionFont="3"
 CaptionSize="10"
 HashSpacing="2.50"
 MarginWidth="1.60"
 LineWidth="0.60"
 BoldWidth="2"
 BondLength="14.40"
 BondSpacing="18"
 ChainAngle="120"
 LabelJustification="Auto"
 CaptionJustification="Left"
 AminoAcidTermini="HOH"
 ShowSequenceTermini="yes"
 ShowSequenceBonds="yes"
 ShowSequenceUnlinkedBranches="no"
 ResidueWrapCount="40"
 ResidueBlockCount="10"
 PrintMargins="36 36 36 36"
 MacPrintInfo="0003000001200120000000000B6608A0FF84FF880BE309180367052703FC0002000001200120000000000B6608A0000100000064000000010001010100000001270F000100010000000000000000000000000002001901900000000000600000000000000000000100000000000000000000000000000000"
 ChemPropName=""
 ChemPropFormula="Chemical Formula: "
 ChemPropExactMass="Exact Mass: "
 ChemPropMolWt="Molecular Weight: "
 ChemPropMOverZ="m/z: "
 ChemPropAnalysis="Elemental Analysis: "
 ChemPropBoilingPt="Boiling Point: "
 ChemPropMeltingPt="Melting Point: "
 ChemPropCritTemp="Critical Temp: "
 ChemPropCritPres="Critical Pres: "
 ChemPropCritVol="Critical Vol: "
 ChemPropGibbs="Gibbs Energy: "
 ChemPropLogP="Log P: "
 ChemPropMR="MR: "
 ChemPropHenry="Henry&apos;s Law: "
 ChemPropEForm="Heat of Form: "
 ChemProptPSA="tPSA: "
 ChemPropCLogP="CLogP: "
 ChemPropCMR="CMR: "
 ChemPropLogS="LogS: "
 ChemPropPKa="pKa: "
 ChemPropID=""
 ChemPropFragmentLabel=""
 color="0"
 bgcolor="1"
 RxnAutonumberStart="1"
 RxnAutonumberConditions="no"
 RxnAutonumberStyle="Roman"
 RxnAutonumberFormat="(#)"
>'''

font_xml = '''<fonttable>
<font id="3" charset="iso-8859-1" name="Arial"/>
</fonttable>'''

cdxml_footer = "</page></CDXML>"

shape_z_counter = 100
text_z_counter = 300
line_z_counter = 1
global_id = 1
# inorder to aviod the same z in generate circle 
already_draw_target= []
# Function to get Bezier curve points

def get_bezier_curve_points_flat(x,y, adjustment_factor=0.05, width_factor=2,adjustment_factor_2=0.1):
    if x == []:
        x_adjusted = ['no_data']
        y = ['no_data']
        bezier_curves = ['no_data']
        return x_adjusted, y, bezier_curves 
    # Calculate the energy change between adjacent energy points (absolute value)
    delta_energy = np.abs(np.diff(y))
    delta_distance = np.abs(np.diff(x))

    # Initialize x coordinates, starting from 0
    x_adjusted = x

    # Adjust x coordinates based on energy changes
    width_factor_adjusted = []
    for i in range(len(x)-1):
        width_factor_adjusted.append(width_factor + delta_distance[i]*delta_distance[i]* adjustment_factor + delta_energy[i]* adjustment_factor_2)  

    x_adjusted = np.array(x_adjusted)
    y = np.array(y)
    width_factor_adjusted = np.array(width_factor_adjusted)

    bezier_curves = []

    # Calculate Bezier control points for each interval
    for i in range(len(x_adjusted) - 1):
        x0, x1 = x_adjusted[i], x_adjusted[i + 1]
        y0, y1 = y[i], y[i + 1]

        # Control points for 6-point Bezier curve: start, end, and two control points for each
        X0, Y0 = x0, y0  # Start point
        X5, Y5 = x1, y1  # End point
        X1, Y1 = x0 - width_factor_adjusted[i], y0  # Control point to the left of start point
        X4, Y4 = x1 + width_factor_adjusted[i], y1  # Control point to the right of end point
        X2, Y2 = x0 + width_factor_adjusted[i], y0  # Control point to the right of start point
        X3, Y3 = x1 - width_factor_adjusted[i], y1  # Control point to the left of end point

        bezier_curves.append({
            'X': [X1, X0, X2, X3, X5, X4],
            'Y': [Y1, Y0, Y2, Y3, Y5, Y4]
        })

    return x_adjusted, y, bezier_curves 

def draw_curve(bezier_curves,connect_type='center',bond_length=10,curves_color = '3',line_type = 'Solid',original_x=None,curve_width=0.6):
    curves_xml = []
    global global_id
    global line_z_counter
    if connect_type == 'center':
        for curve_id, curve in enumerate(bezier_curves, start=1):
            X = curve['X']
            Y = curve['Y']
            scaled_points = [f"{X[i]:.2f} {Y[i]:.2f}" for i in range(len(X))]  # Flip Y coordinates
            curve_points_str = " ".join(scaled_points)
            curves_xml.append(f'<curve id="{global_id}"\n Z="{line_z_counter}\n" color="{curves_color}"\n LineType="{line_type}"\n LineWidth="{curve_width}"\n CurvePoints="{curve_points_str}"\n />')
            global_id += 1
            line_z_counter += 1
    elif connect_type == 'side':
        for curve_id, curve in enumerate(bezier_curves, start=1):
            X = curve['X']
            Y = curve['Y']

            X[0] = X[0] + bond_length*(2*original_x[curve_id-1]+3)
            X[1] = X[1] + bond_length*(2*original_x[curve_id-1]+3)
            X[2] = X[2] + bond_length*(2*original_x[curve_id-1]+3)

            X[3] = X[3] + bond_length*(2*original_x[curve_id]+1)
            X[4] = X[4] + bond_length*(2*original_x[curve_id]+1)
            X[5] = X[5] + bond_length*(2*original_x[curve_id]+1)

            scaled_points = [f"{X[i] :.2f} {Y[i]:.2f}" for i in range(len(X))]  # Flip Y coordinates
            curve_points_str = " ".join(scaled_points)
            curves_xml.append(f'<curve id="{global_id}"\n Z="{line_z_counter}"\n color="{curves_color}"\n LineType="{line_type}"\n LineWidth="{curve_width}"\n CurvePoints="{curve_points_str}"\n />')
            global_id += 1
            line_z_counter += 1
    return curves_xml

def draw_line(center_x,center_y,linetype,linewidth,bond_length=10,bond_color = 4,connect_type = 'side',original_x=None):
    line_xml = []
    global global_id
    global line_z_counter


    if connect_type == 'center':
        # Transform center coordinates
        scale_center_x = center_x
        scale_center_y = center_y
    elif connect_type == 'side':
        # Transform center coordinates
        scale_center_x = [x + bond_length*2*(original_x[idx]+1) for idx, x in enumerate(center_x)]
        scale_center_y = center_y

    for i in range(len(scale_center_x)-1):
        # Calculate bounding box coordinates
        if connect_type == 'center':
            x1 = scale_center_x[i] 
            x2 = scale_center_x[i+1]
            y1 = scale_center_y[i]
            y2 = scale_center_y[i+1]
            bounding_box = f"{x1:.2f}  {y1:.2f} {x2:.2f} {y2:.2f}"
        elif connect_type == 'side':
            x1 = scale_center_x[i]+ bond_length
            x2 = scale_center_x[i+1] - bond_length
            y1 = scale_center_y[i]
            y2 = scale_center_y[i+1]
            bounding_box = f"{x1:.2f}  {y1:.2f} {x2:.2f} {y2:.2f}"

        line = ET.Element('arrow', id=str(global_id), Z=str(line_z_counter), LineType=f'{linetype} Bold')
        global_id += 1
        line_z_counter += 1
        # Set BoundingBox attribute
        line.set("BoundingBox", bounding_box)

        # Set color attribute (if provided)
        if bond_color:
            line.set("color", str(bond_color))

                # Set 3D coordinates
        head3d = f"{x2:.2f} {scale_center_y[i+1] :.2f} 0"
        tail3d = f"{x1:.2f} {scale_center_y[i] :.2f} 0"
        center3d = f"{scale_center_x[i] :.2f} {scale_center_y[i]:.2f} 0"
        major_axis_end3d = f"{x2:.2f} {y2:.2f} 0"
        minor_axis_end3d = f"{x1:.2f} {y2:.2f} 0"

        line.set("BoldWidth", str(linewidth))
        line.set("Head3D", head3d)
        line.set("Tail3D", tail3d)
        line.set("Center3D", center3d)
        line.set("MajorAxisEnd3D", major_axis_end3d)
        line.set("MinorAxisEnd3D", minor_axis_end3d)

        # Convert to string and append to list
        line_xml.append(ET.tostring(line, encoding='unicode', method='xml'))
    
    return line_xml

def generate_rectangle_xml(center_x,center_y,bond_length=10,bond_color = 4,connect_type = 'side',original_x=None,bond_width = 2.0):
    rectangles_xml = []
    global global_id
    global shape_z_counter
    global page_width
    global page_high
    global already_draw_target

    if connect_type == 'center':
        # Transform center coordinates
        scale_center_x = center_x
        scale_center_y = center_y
    elif connect_type == 'side':
        # Transform center coordinates
        scale_center_x = [x + bond_length*2*(original_x[idx]+1) for idx, x in enumerate(center_x)]
        scale_center_y = center_y

    # auto change page width and high
    for x in  scale_center_x:
        ad_page_width = math.ceil(abs(x/510))
        if ad_page_width > page_width:
            page_width = ad_page_width
    for y in scale_center_y:
        ad_page_high = math.ceil(abs(y/710))
        if ad_page_high > page_high:
            page_high = ad_page_high


    for i in range(len(scale_center_x)):
        coordinate = (scale_center_x[i], scale_center_y[i])
        if coordinate not in already_draw_target:
            already_draw_target.append(coordinate)
        else:
            continue
        # Calculate bounding box coordinates
        x1 = (scale_center_x[i] - bond_length) 
        x2 = (scale_center_x[i] + bond_length)
        y1 = scale_center_y[i]
        y2 = scale_center_y[i]
        bounding_box = f"{x1:.2f} {y1:.2f} {x2:.2f} {y2:.2f}"
        
        arrow = ET.Element('arrow', id=str(global_id), Z=str(shape_z_counter), LineType="Bold", FillType="None", ArrowheadType="Solid")
        global_id += 1
        shape_z_counter += 1
        # Set BoundingBox attribute
        arrow.set("BoundingBox", bounding_box)

        # Set color attribute (if provided)
        if bond_color:
            arrow.set("color", str(bond_color))

        # Set 3D coordinates
        head3d = f"{x2:.2f} {scale_center_y[i] :.2f} 0"
        tail3d = f"{x1:.2f} {scale_center_y[i] :.2f} 0"
        center3d = f"{scale_center_x[i] :.2f} {scale_center_y[i]:.2f} 0"
        major_axis_end3d = f"{x2:.2f} {y2:.2f} 0"
        minor_axis_end3d = f"{x1:.2f} {y2:.2f} 0"

        arrow.set("BoldWidth", str(bond_width))
        arrow.set("Head3D", head3d)
        arrow.set("Tail3D", tail3d)
        arrow.set("Center3D", center3d)
        arrow.set("MajorAxisEnd3D", major_axis_end3d)
        arrow.set("MinorAxisEnd3D", minor_axis_end3d)
        
        # Convert to string and append to list
        rectangles_xml.append(ET.tostring(arrow, encoding='unicode', method='xml'))
    
    return rectangles_xml

def generate_circle_xml(center_x, center_y, radius=5, circle_color=8, connect_type='side',original_x=None):
    import xml.etree.ElementTree as ET
    global shape_z_counter
    global global_id
    global page_width
    global page_high
    global already_draw_target
    
    circles_xml = []

    if connect_type == 'center':
        # Transform center coordinates
        scale_center_x = center_x
        scale_center_y = center_y
    elif connect_type == 'side':
        # Transform center coordinates
        scale_center_x = [x + radius *2*(original_x[idx]+1) for idx, x in enumerate(center_x)]
        scale_center_y = center_y

    # auto change page width and high
    for x in  scale_center_x:
        ad_page_width = math.ceil(abs(x/500))
        if ad_page_width > page_width:
            page_width = ad_page_width
    for y in scale_center_y:
        ad_page_high = math.ceil(abs(y/710))
        if ad_page_high > page_high:
            page_high = ad_page_high

    for i in range(len(scale_center_x)):
        coordinate = (scale_center_x[i], scale_center_y[i])
        if coordinate not in already_draw_target:
            already_draw_target.append(coordinate)
        else:
            continue
        # Calculate bounding box coordinates
        x1 = (scale_center_x[i] - radius)
        x2 = (scale_center_x[i] + radius)
        y1 = (scale_center_y[i] - radius)
        y2 = (scale_center_y[i] + radius)
        bounding_box = f"{x1:.2f} {y1:.2f} {x2:.2f} {y2:.2f}"

        # Create the root element
        circle_id = 20 + i  # Example id, can be parameterized
        graphic = ET.Element('graphic', id=str(global_id), Z=str(shape_z_counter), color=str(circle_color), GraphicType="Oval", OvalType="Circle Filled")
        shape_z_counter += 1
        global_id += 1
        # Set BoundingBox attribute
        graphic.set("BoundingBox", bounding_box)

        # Set 3D coordinates
        center3d = f"{scale_center_x[i]:.2f} {scale_center_y[i]:.2f} 0"
        major_axis_end3d = f"{x2:.2f} {scale_center_y[i]:.2f} 0"
        minor_axis_end3d = f"{scale_center_x[i]:.2f} {y2:.2f} 0"

        graphic.set("Center3D", center3d)
        graphic.set("MajorAxisEnd3D", major_axis_end3d)
        graphic.set("MinorAxisEnd3D", minor_axis_end3d)

        # Convert to string and append to list
        circles_xml.append(ET.tostring(graphic, encoding='unicode', method='xml'))

    return circles_xml

def generate_text_cdxml(center_x, center_y, energy_text_list, target_text_list,location,target_layout,target_location,text_space,text_base_movement,target_move,target_move_y, text_color=4, connect_type='side', original_x=None, font_size=10, font_type=3, Z_value=70):

    """
    Generates a list of cdxml text elements with a given set of coordinates and text.
    
    Args:
        center_x (list): List of x coordinates for each text element.
        center_y (list): List of y coordinates for each text element.
        energy_text_list (list): List of energy text
        target_text_list (list): List of target_text
        locaton (list): List of text Location. sc sw sa ss sd cc cw cs ca cd 
        target_layout (str): global site of taget location.combine location
        target_location (str): global site of taget location. c w s a d
        text_space (float): text_spcae to marker
        text_base_movement (float): move test to mark center
        target_move (float): move caused by tagert size (bond_length/radius)
        text_color (int, optional): Color code for the text (default is 4).
        font_size (int, optional): Font size for the text (default is 10).
        font_type (int, optional): Font type for the text (default is 3).
        Z_value (int, optional): Z-value for the <t> tag (default is 70).
        connect_type (str, optional): Connection type, either 'center' or 'side' (default is 'side').
        original_x (list, optional): List of original x values, used when connect_type is 'side'.
        
    Returns:
        list: List of cdxml text elements in string format.
    """
    
    cdxml_elements = []
    def add_text_xml(p_x,p_y, text, text_color, font_size, font_type, Z_value, cdxml_elements, i):
        global global_id
        global text_z_counter
        if text != '':
            text = str(text)

                # if len(text) == 1 and text.isalpha() and text.isupper():
                #     p_x -= 3

            # Create the <t> element with the corresponding attributes
            t_element = ET.Element('t', id=str(global_id), p=f"{p_x:.2f} {p_y:.2f}", Z=str(text_z_counter),CaptionJustification="Center",Justification="Center",LineHeight="auto",InterpretChemically="no")
            global_id += 1
            text_z_counter += 1
            # Create the <s> element for text styling
            s_element = ET.SubElement(t_element, 's', font=str(font_type), size=str(font_size), color=str(text_color))
            s_element.text = text
                
                # Convert the element to string and append to the list
            cdxml_elements.append(ET.tostring(t_element, encoding='unicode', method='xml'))
        
    # Adjust coordinates based on connect_type
    if connect_type == 'center':
        # No transformation, use provided coordinates directly
        scale_center_x = center_x
        scale_center_y = [y + text_base_movement  for y in center_y]
    elif connect_type == 'side' and original_x is not None:
        # Adjust x-coordinates based on original_x and radius (as in your previous example)
        scale_center_x = [x +  target_move * 2 *(original_x[idx]+1) for idx, x in enumerate(center_x)]
        scale_center_y = [y + text_base_movement  for y in center_y]
    else:
        raise ValueError("Invalid 'connect_type' or missing 'original_x' for 'side'.")
    
    target_move_y += font_size*0.5

    for idx in range (len(scale_center_x)):
        i = idx
        coordinate = (scale_center_x[i], scale_center_y[i])
        if coordinate not in already_draw_target:
            already_draw_target.append(coordinate)
        else:
            continue
        if location[idx].lower() == 'sc':
            add_text_xml(scale_center_x[idx],scale_center_y[idx]-text_space-target_move_y, energy_text_list[idx], text_color, font_size, font_type, Z_value, cdxml_elements, i)
            if target_text_list[idx]:
                add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+target_move_y,target_text_list[idx],text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'cw':
            if target_text_list[idx]:
                combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
            else:
                combine_text = str(energy_text_list[idx])
            add_text_xml(scale_center_x[idx], scale_center_y[idx]-text_space-target_move_y,combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'cs':
            if target_text_list[idx]:
                combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
            else:
                combine_text = str(energy_text_list[idx])
            add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+target_move_y,combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'ca':
            if target_text_list[idx]:
                combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
            else:
                combine_text = str(energy_text_list[idx])
            add_text_xml(scale_center_x[idx]-text_space-target_move-len(combine_text)/2*font_size*0.5, scale_center_y[idx],combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'cd':
            if target_text_list[idx]:
                combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
            else:
                combine_text = str(energy_text_list[idx])
            add_text_xml(scale_center_x[idx]+text_space+target_move+len(combine_text)/2*font_size*0.5, scale_center_y[idx],combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'cc':
            if target_text_list[idx]:
                combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
            else:
                combine_text = str(energy_text_list[idx]) 
            add_text_xml(scale_center_x[idx], scale_center_y[idx],combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)

        elif location[idx].lower() == 'sw':
            add_text_xml(scale_center_x[idx], scale_center_y[idx]-text_space-target_move_y,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
            if target_text_list[idx]:
                add_text_xml(scale_center_x[idx], scale_center_y[idx]-text_space-font_size-target_move_y,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'ss':
            add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+target_move_y,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
            if target_text_list[idx]:
                add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+font_size+target_move_y,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'sa':
            if target_text_list[idx]:
                add_text_xml(scale_center_x[idx]-text_space-target_move-len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]-font_size/2,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                add_text_xml(scale_center_x[idx]-text_space-target_move-len(str(target_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]+font_size/2,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
            else:
                add_text_xml(scale_center_x[idx]-text_space-target_move-len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx],str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
        elif location[idx].lower() == 'sd':
            if target_text_list[idx]:
                add_text_xml(scale_center_x[idx]+text_space+target_move+len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]-font_size/2,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                add_text_xml(scale_center_x[idx]+text_space+target_move+len(str(target_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]+font_size/2,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
            else:
                add_text_xml(scale_center_x[idx]+text_space+target_move+len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx],str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
        else:
            if target_layout == 'sperate':
                if target_location == 'c': #sc
                    add_text_xml(scale_center_x[idx],scale_center_y[idx]-text_space-target_move_y, energy_text_list[idx], text_color, font_size, font_type, Z_value, cdxml_elements, i)
                    if target_text_list[idx]:
                        add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+target_move_y,target_text_list[idx],text_color, font_size, font_type, Z_value, cdxml_elements, i)

                elif target_location == 'w': #sw
                    add_text_xml(scale_center_x[idx], scale_center_y[idx]-text_space-target_move_y,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                    if target_text_list[idx]:
                        add_text_xml(scale_center_x[idx], scale_center_y[idx]-text_space-font_size-target_move_y,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)

                elif target_location == 's': #ss
                    add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+target_move_y,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                    if target_text_list[idx]:
                        add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+font_size+target_move_y,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)   

                elif target_location == 'a': #sa
                    if target_text_list[idx]:
                        add_text_xml(scale_center_x[idx]-text_space-target_move-len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]-font_size/2,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                        add_text_xml(scale_center_x[idx]-text_space-target_move-len(str(target_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]+font_size/2,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                    else:
                        add_text_xml(scale_center_x[idx]-text_space-target_move-len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx],str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                elif target_location == 'd': #sd
                    if target_text_list[idx]:
                        add_text_xml(scale_center_x[idx]+text_space+target_move+len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]-font_size/2,str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                        add_text_xml(scale_center_x[idx]+text_space+target_move+len(str(target_text_list[idx]))/2*font_size*0.5, scale_center_y[idx]+font_size/2,str(target_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)
                    else:
                        add_text_xml(scale_center_x[idx]+text_space+target_move+len(str(energy_text_list[idx]))/2*font_size*0.5, scale_center_y[idx],str(energy_text_list[idx]),text_color, font_size, font_type, Z_value, cdxml_elements, i)

            elif target_layout == 'combine':
                if target_location == 'c': #cc
                    if target_text_list[idx]:
                        combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
                    else:
                        combine_text = str(energy_text_list[idx]) 
                    add_text_xml(scale_center_x[idx], scale_center_y[idx],combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)

                elif  target_location == 'w': #cw
                    if target_text_list[idx]:
                        combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
                    else:
                        combine_text = str(energy_text_list[idx])
                    add_text_xml(scale_center_x[idx], scale_center_y[idx]-text_space-target_move_y,combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)

                elif  target_location == 's': #cs 
                    if target_text_list[idx]:
                        combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
                    else:
                        combine_text = str(energy_text_list[idx])
                    add_text_xml(scale_center_x[idx], scale_center_y[idx]+text_space+target_move_y,combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)

                elif  target_location == 'a': #ca
                    if target_text_list[idx]:
                        combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
                    else:
                        combine_text = str(energy_text_list[idx])
                    add_text_xml(scale_center_x[idx]-text_space-target_move-len(combine_text)/2*font_size*0.5, scale_center_y[idx],combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)

                elif  target_location == 'd': #cd     
                    if target_text_list[idx]:
                        combine_text = f'{energy_text_list[idx]} {target_text_list[idx]}'
                    else:
                        combine_text = str(energy_text_list[idx])
                    add_text_xml(scale_center_x[idx]+text_space+target_move+len(combine_text)/2*font_size*0.5, scale_center_y[idx],combine_text,text_color, font_size, font_type, Z_value, cdxml_elements, i)
                        
    return cdxml_elements

# Function to save CDXML to file
def save_cdxml_file(cdxml_string, filename="output.cdxml"):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(cdxml_string)

def adjust_line_width_based_on_chart_ratio(ax, scaled_x_adjusted, scaled_y_points):
    """
    通过图表的宽高比例和scaled_x_adjusted, scaled_y_points中最大差值，确定是绘图的坐标轴依据哪个变化的，
    并根据此调整线宽。
    
    参数：
        ax: 当前的matplotlib轴 (ax) 对象
        scaled_x_adjusted: x轴上的调整坐标列表
        scaled_y_points: y轴上的调整坐标列表
        line_width: 默认线宽
    """
    line_width = 30
    
    # 计算scaled_x_adjusted和scaled_y_points中的最大差值
    x_diff = max(scaled_x_adjusted) - min(scaled_x_adjusted)
    y_diff = max(scaled_y_points) - min(scaled_y_points)

    # 获取图表的位置和大小比例 [left, bottom, width, height]
    fig_size = fig.get_size_inches()
    chart_width = fig_size[0]
    chart_height = fig_size[1]


    # 计算图表的比例，比较 x 和 y 轴的变化
    if x_diff/chart_width > y_diff / chart_height:
        # 如果 x 轴的变化范围占比更大，依据 x 轴调整线宽
        ratio = chart_width / x_diff
        adjusted_line_width = line_width * ratio
    else:
        # 如果 y 轴的变化范围占比更大，依据 y 轴调整线宽
        ratio = chart_height / y_diff
        adjusted_line_width = line_width * ratio

    # 返回调整后的线宽
    return adjusted_line_width

hlines = []
curve_lines = []
target_texts = []
energy_texts= []
combine_texxts = []
size_markers = []

def interactive_bezier_curve():

    def parse_data(data):
        # 去除两端的空白字符
        data = data.strip()

        # 初始化默认值
        y_value = None
        target = ""
        location = "bs"  # 默认 location 值为 "bs"

        # 替换中文括号和中文逗号为英文括号和逗号
        data = data.replace('（', '(').replace('）', ')').replace('，', ',')


        if ',' in data:
            y_split = data.split(',')
            y_and_target = y_split[0]
            if  y_split[1].strip():
                location = y_split[1].strip()
        else:
            y_and_target = data 


        if '(' in y_and_target:
            y_value =  y_and_target.split('(')[0]
            target = y_and_target.split('(')[1].split(')')[0].strip()

        else:
            y_value = y_and_target

    
        return [y_value, target, location]
    
    def draw_benzene(ax, center_x, center_y, line_width,hexagon_side, bond_gap, bond_length_ratio, ):
        """
        Draw a benzene-like hexagon with shorter alternating double bonds on a matplotlib Axes.

        Args:
            ax (matplotlib.axes.Axes): The Axes to draw on.
            center_x (float): X-coordinate of the benzene center.
            center_y (float): Y-coordinate of the benzene center.
            hexagon_side (float): Length of each side of the hexagon.
            bond_gap (float): Gap between the lines of the double bond.
            bond_length_ratio (float): Proportion of the bond length relative to the full side.
            line_width (float): Width of the hexagon and double bonds.
            size_markers (list): List to store each plotted line as an element.
        """
        # Calculate the coordinates for the hexagon vertices
        angles = np.linspace(0, 2 * np.pi, 7)  # 6 vertices + closing point
        x = center_x + hexagon_side * np.cos(angles)
        y = center_y + hexagon_side * np.sin(angles)

        # Plot the hexagon
        hexagon_line, = ax.plot(x, y, 'k-', linewidth=line_width)  # Extract the Line2D object
        size_markers.append(hexagon_line)

        # Double bond indices
        bond_indices = [(0, 1), (2, 3), (4, 5)]  # Pairs of vertices for double bonds
        for i, j in bond_indices:
            # Calculate the direction vector of the bond
            dx = x[j] - x[i]
            dy = y[j] - y[i]
            length = np.sqrt(dx**2 + dy**2)
            dx /= length
            dy /= length

            # Calculate the shorter bond start and end points
            x_start = x[i] + (1 - bond_length_ratio) / 2 * length * dx
            y_start = y[i] + (1 - bond_length_ratio) / 2 * length * dy
            x_end = x[j] - (1 - bond_length_ratio) / 2 * length * dx
            y_end = y[j] - (1 - bond_length_ratio) / 2 * length * dy

            # Perpendicular offset for the second line
            perp_dx = dy  # Rotate 90 degrees
            perp_dy = -dx

            # Plot the shorter double bond as two close lines
            double_bond_line, = ax.plot([x_start - bond_gap * perp_dx, x_end - bond_gap * perp_dx], 
                                        [y_start - bond_gap * perp_dy, y_end - bond_gap * perp_dy], 
                                        'k-', linewidth=line_width)
            size_markers.append(double_bond_line)

    def update_plot(*args):

        try:
            # Get user input parameters
            adjustment_factor = float(adjustment_factor_var.get())
            width_factor = float(width_factor_var.get())
            adjustment_factor_2 = float(adjustment_factor_2_var.get())
            scale_factor = float(scale_factor_var.get())
            scale_factor_x = float(scale_factor_x_var.get())
            scale_factor_y = float(scale_factor_y_var.get())
            bond_length = float(bond_length_var.get())
            radius = float(radius_var.get())
            connect_type = connect_type_var.get()
            shape_type = shape_type_var.get()
            line_type = line_type_var.get()
            font_size = font_size_var.get()
            text_space = text_space_var.get()
            show_marker = show_marker_var.get()
            target_layout = target_layout_var.get()
            target_location = target_location_var.get()
            curve_width = curve_width_var.get()
            bond_width = bond_width_var.get()
            grid = grid_var.get()

            if table is None or not table.winfo_exists():
                print('No table window detected, please click "Open Table"')
                return

            global already_draw_target
            already_draw_target = []
            

            x_min_previous, x_max_previous = ax.get_xlim()
            y_min_previous, y_max_previous = ax.get_ylim()

            x_range_previous = abs(x_max_previous - x_min_previous)
            y_range_previous = abs(y_max_previous - y_min_previous)

            if x_range_previous >10 and y_range_previous >10:
                contain_ax = True
            else:
                contain_ax = False
            
            # Clear previous plot
            ax.clear()
            ax.tick_params(axis='both', which='both', bottom=False, top=False,
                left=False, right=False, labelbottom=False, labelleft=False)
            # ax.set_title("Interactive Bezier Curve with Shapes")
            # ax.set_xlabel("Adjusted Coordinate")
            # ax.set_ylabel("Potential Energy")
            if grid:
                ax.grid(True)

            # Invert Y-axis
            ax.invert_yaxis()

            # Set equal aspect ratio to maintain shape consistency
            ax.set_aspect("equal", adjustable="datalim")

            # line_width = adjust_line_width_based_on_chart_ratio(ax, x_total, y_total)
            line_width = 10
            key_width =  line_width*3.38

            # Get data from the table and create y lists for each row
            for row in table.get_children():
                values = table.item(row)['values']
                original_y = []
                original_x = []
                target = []
                location = []
                for index, value in enumerate(values[3:]):
                    value = str(value)
                    if str(value).strip():  # If value is not empty
                        try:
                            value = parse_data(value)
                            original_y.append(float(value[0]))
                            target.append(value[1])
                            location.append(value[2])
                            original_x.append(index + 1)
                        except ValueError:
                            print(f"Invalid value at row {row}, column {index + 2}. Please correct it.")
                            return

                curve_color = values[0]
                marker_color = values[1]
                text_color = values[2]

                # Call Bezier curve calculation function (dummy function here)
                x_adjusted, y_points, bezier_curves = get_bezier_curve_points_flat(
                original_x, original_y, adjustment_factor, width_factor, adjustment_factor_2
                )
                
                if  x_adjusted[0] == 'no_data': # no data line
                    continue

                # Scale Bezier curves
                for curve in bezier_curves:
                    X = curve['X']
                    Y = curve['Y']
                    for i in range(len(X)):
                        X[i] = X[i] * scale_factor *scale_factor_x* 10 + 50
                        Y[i] = 500 - Y[i] * scale_factor*scale_factor_y / 2

                # Scale adjusted coordinates and potential energy values
                scaled_x_adjusted = [x * scale_factor*scale_factor_x * 10 + 50 for x in x_adjusted]
                scaled_y_points = [500 - y * scale_factor*scale_factor_y / 2 for y in y_points]

                fontsize = line_width*15 # 随便设置的大小

                # Draw Bezier curves based on connect_type
                for i, curve in enumerate(bezier_curves):
                    X = curve['X']
                    Y = curve['Y']
                    t = np.linspace(0, 1, 200)

                    if connect_type == "center":
                        # Continuous curve
                        Bx = (1 - t)**3 * X[1] + 3 * (1 - t)**2 * t * X[2] + 3 * (1 - t) * t**2 * X[3] + t**3 * X[4]
                        By = (1 - t)**3 * Y[1] + 3 * (1 - t)**2 * t * Y[2] + 3 * (1 - t) * t**2 * Y[3] + t**3 * Y[4]
                        if line_type == "Dashed":
                            curve_line = ax.plot(Bx, By, color=curve_color,linewidth=line_width,linestyle=line_type.lower(),dashes=(5*0.6/curve_width, 5*0.6/curve_width))
                            curve_lines.append(curve_line)
                        elif line_type == "Solid":
                            curve_line = ax.plot(Bx, By, color=curve_color,linewidth=line_width,linestyle=line_type.lower())
                            curve_lines.append(curve_line)

                    elif connect_type == "side":
                        # Non-continuous curve with adjusted endpoints
                        if shape_type == "line":
                            offset = bond_length 
                        elif shape_type == "circle":
                            offset = radius  
                        else:
                            offset = 0

                        X[0] = X[0] + offset*(2*original_x[i]+1)
                        X[1] = X[1] + offset*(2*original_x[i]+1)
                        X[2] = X[2] + offset*(2*original_x[i]+1)

                        X[3] = X[3] + offset*(2*original_x[i+1]-1)
                        X[4] = X[4] + offset*(2*original_x[i+1]-1)
                        X[5] = X[5] + offset*(2*original_x[i+1]-1)

                            
                        Bx = (1 - t)**3 * X[1] + 3 * (1 - t)**2 * t * X[2] + 3 * (1 - t) * t**2 * X[3] + t**3 * X[4]
                        By = (1 - t)**3 * Y[1] + 3 * (1 - t)**2 * t * Y[2] + 3 * (1 - t) * t**2 * Y[3] + t**3 * Y[4]
                        if line_type == "Dashed":
                            curve_line = ax.plot(Bx, By, color=curve_color,linewidth=line_width,linestyle=line_type.lower(),dashes=(5*0.6/curve_width, 5*0.6/curve_width))
                            curve_lines.append(curve_line)
                        elif line_type == "Solid":
                            curve_line = ax.plot(Bx, By, color=curve_color,linewidth=line_width,linestyle=line_type.lower())
                            curve_lines.append(curve_line)

                # Draw horizontal lines or circles
                if shape_type == "line":
                    text_shape_space = bond_length
                    text_shape_space_y = bond_width/2 + font_size*0.5
                    if connect_type == "side":
                        for idx, (x, y) in enumerate(zip(scaled_x_adjusted, scaled_y_points)):
                            center_x = x + bond_length*2*(original_x[idx])
                            coordinate = (center_x, y)
                            if coordinate not in already_draw_target:
                                already_draw_target.append(coordinate)
                            else:
                                continue
                            print_text(font_size, text_space, original_y, target, location, text_color, fontsize, text_shape_space, text_shape_space_y,idx, y, center_x,target_layout,target_location)
                            hline = ax.hlines(y, center_x - bond_length, center_x + bond_length, colors=marker_color, linewidth=key_width,zorder=100)
                            hlines.append(hline)
                            
                    elif connect_type == "center":
                        for idx, (x, y) in enumerate(zip(scaled_x_adjusted, scaled_y_points)):
                            coordinate = (x, y)
                            if coordinate not in already_draw_target:
                                already_draw_target.append(coordinate)
                            else:
                                continue
                            print_text(font_size, text_space, original_y, target, location, text_color, fontsize, text_shape_space,text_shape_space_y, idx, y, x,target_layout,target_location)
                            hline = ax.hlines(y, x - bond_length, x + bond_length, colors=marker_color, linewidth=key_width,zorder=100)
                            hlines.append(hline)
                            
    
                elif shape_type == "circle":
                    text_shape_space = radius
                    text_shape_space_y = radius + font_size*0.5
                    if connect_type == "side":
                        for idx, (x, y) in enumerate(zip(scaled_x_adjusted, scaled_y_points)):
                            center_x = x + radius*2*(original_x[idx])
                            coordinate = (center_x, y)
                            if coordinate not in already_draw_target:
                                already_draw_target.append(coordinate)
                            else:
                                continue
                            print_text(font_size, text_space, original_y, target, location, text_color, fontsize, text_shape_space, text_shape_space_y,idx, y, center_x,target_layout,target_location)
                            circle = plt.Circle((center_x, y), radius, color=marker_color, fill=True,zorder=100)
                            ax.add_artist(circle)
                            
                    elif connect_type == "center":
                        for idx, (x, y) in enumerate(zip(scaled_x_adjusted, scaled_y_points)):
                            coordinate = (x, y)
                            if coordinate not in already_draw_target:
                                already_draw_target.append(coordinate)
                            else:
                                continue
                            print_text(font_size, text_space, original_y, target, location, text_color, fontsize, text_shape_space, text_shape_space_y,idx, y, x,target_layout,target_location)
                            circle = plt.Circle((x, y), radius, color=marker_color, fill=True,zorder=100)
                            ax.add_artist(circle)
                            


                elif shape_type == "None":
                    text_shape_space = 0
                    text_shape_space_y = 0 + font_size*0.5
                    for idx, (x, y) in enumerate(zip(scaled_x_adjusted, scaled_y_points)):
                        coordinate = (x, y)
                        if coordinate not in already_draw_target:
                            already_draw_target.append(coordinate)
                        else:
                            continue
                        print_text(font_size, text_space, original_y, target, location, text_color, fontsize, text_shape_space,text_shape_space_y, idx, y, x,target_layout,target_location)


            
            canvas.draw()

            if contain_ax == True:
                ax.set_xlim(x_min_previous, x_max_previous)
                ax.set_ylim(y_min_previous, y_max_previous)

            x_min, x_max = ax.get_xlim()
            y_min, y_max = ax.get_ylim()
            center_x = (x_min+x_max)/2
            center_y = (y_min+y_max)/2

            # 缩放因子计算
            x_range = x_max - x_min
            y_range = y_min - y_max
            width, height = canvas.get_width_height()

            fontsize = font_size/max(x_range/width, y_range/height)*0.7
            hline_linewidth = bond_width/max(x_range/width, y_range/height)*0.7
            curve_linewidth = curve_width/max(x_range/width, y_range/height)*0.7
            size_marker_linewidth = 0.6/max(x_range/width, y_range/height)*0.7


            if contain_ax == True:
                ax.set_xlim(x_min_previous, x_max_previous)
                ax.set_ylim(y_min_previous, y_max_previous)

            for text_obj in ax.texts:
                # current_fontsize = text_obj.get_fontsize()
                # print(current_fontsize)
                text_obj.set_fontsize(fontsize)

            for hline_obj in hlines:
                hline_obj.set_linewidth(hline_linewidth)

            for curve_obj in curve_lines:
                for curve in curve_obj:
                    curve.set_linewidth(curve_linewidth)
            
            #draw maker
            if show_marker:
                draw_benzene(ax, center_x, center_y, size_marker_linewidth,hexagon_side=14, bond_gap=2.5, bond_length_ratio=0.7)

            # 重新绘制画布
            canvas.draw()

                
        except ValueError:
            tk.messagebox.showerror("Input Error", "Please enter valid numeric values.")

    def print_text(font_size, text_space, original_y, target, location, text_color, fontsize, text_shape_space,text_shape_space_y, idx, y, center_x,target_layout,target_location):
        if location[idx].lower() == 'sc':
            energy_text = ax.text(center_x, y-text_space-text_shape_space_y, str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
            if target[idx]:
                target_text = ax.text(center_x, y+text_space+text_shape_space_y , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'cw':
            if target[idx]:
                combine_text = f'{original_y[idx]} {target[idx]}'
            else:
                combine_text = str(original_y[idx])
            energy_text = ax.text(center_x, y-text_space-text_shape_space_y, combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'cs':
            if target[idx]:
                combine_text = f'{original_y[idx]} {target[idx]}'
            else:
                combine_text = str(original_y[idx])
            energy_text = ax.text(center_x, y+text_space+text_shape_space_y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'ca':
            if target[idx]:
                combine_text = f'{original_y[idx]} {target[idx]}'
            else:
                combine_text = str(original_y[idx])
            energy_text = ax.text(center_x-text_space-text_shape_space-len(combine_text)/2*font_size*0.5, y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'cd':
            if target[idx]:
                combine_text = f'{original_y[idx]} {target[idx]}'
            else:
                combine_text = str(original_y[idx])
            energy_text = ax.text(center_x+text_space+text_shape_space+len(combine_text)/2*font_size*0.5, y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'cc':
            if target[idx]:
                combine_text = f'{original_y[idx]} {target[idx]}'
            else:
                combine_text = str(original_y[idx]) 
            energy_text = ax.text(center_x, y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})  

        elif location[idx].lower() == 'sw':
            energy_text = ax.text(center_x, y-text_space-text_shape_space_y, str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
            if target[idx]:
                target_text = ax.text(center_x, y-text_space-font_size-text_shape_space_y, str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'ss':
            energy_text = ax.text(center_x, y+text_space+text_shape_space_y , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
            if target[idx]:
                target_text = ax.text(center_x, y+text_space+font_size+text_shape_space_y, str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'sa':
            if target[idx]:
                energy_text = ax.text(center_x-text_space-text_shape_space-len(str(original_y[idx]))/2*font_size*0.5, y-font_size/2 , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                target_text = ax.text(center_x-text_space-text_shape_space-len(str(target[idx]))/2*font_size*0.5, y+font_size/2 , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
            else:
                energy_text = ax.text(center_x-text_space-text_shape_space-len(str(original_y[idx]))/2*font_size*0.5, y , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        elif location[idx].lower() == 'sd':
            if target[idx]:
                energy_text = ax.text(center_x+text_space+text_shape_space+len(str(original_y[idx]))/2*font_size*0.5, y-font_size/2 , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                target_text = ax.text(center_x+text_space+text_shape_space+len(str(target[idx]))/2*font_size*0.5, y+font_size/2 , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
            else:
                energy_text = ax.text(center_x+text_space+text_shape_space+len(str(original_y[idx]))/2*font_size*0.5, y , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
        
        else:
            if target_layout == 'sperate':
                if target_location == 'c': #sc
                    energy_text = ax.text(center_x, y-text_space-text_shape_space_y , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                    energy_texts.append(energy_text)
                    if target[idx]:
                        target_text = ax.text(center_x, y+text_space+text_shape_space_y, str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                        target_texts.append(target_text)
                elif target_location == 'w': #sw
                    energy_text = ax.text(center_x, y-text_space-text_shape_space_y  , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                    if target[idx]:
                        target_text = ax.text(center_x, y-text_space-font_size-text_shape_space_y , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                elif target_location == 's': #ss   
                    energy_text = ax.text(center_x, y+text_space+text_shape_space_y  , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                    if target[idx]:
                        target_text = ax.text(center_x, y+text_space+font_size+text_shape_space_y  , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                elif target_location == 'a': #sa
                    if target[idx]:
                        energy_text = ax.text(center_x-text_space-text_shape_space-len(str(original_y[idx]))/2*font_size*0.5, y-font_size/2 , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                        target_text = ax.text(center_x-text_space-text_shape_space-len(str(target[idx]))/2*font_size*0.5, y+font_size/2 , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                    else:
                        energy_text = ax.text(center_x-text_space-text_shape_space-len(str(original_y[idx]))/2*font_size*0.5, y , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                elif target_location == 'd': #sd
                    if target[idx]:
                        energy_text = ax.text(center_x+text_space+text_shape_space+len(str(original_y[idx]))/2*font_size*0.5, y-font_size/2 , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                        target_text = ax.text(center_x+text_space+text_shape_space+len(str(target[idx]))/2*font_size*0.5, y+font_size/2 , str(target[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                    else:
                        energy_text = ax.text(center_x+text_space+text_shape_space+len(str(original_y[idx]))/2*font_size*0.5, y , str(original_y[idx]), color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                
            elif target_layout == 'combine':
                if target_location == 'c': #cc
                    if target[idx]:
                        combine_text = f'{original_y[idx]} {target[idx]}'
                    else:
                        combine_text = str(original_y[idx]) 
                    energy_text = ax.text(center_x, y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})  

                elif  target_location == 'w': #cw
                    if target[idx]:
                        combine_text = f'{original_y[idx]} {target[idx]}'
                    else:
                        combine_text = str(original_y[idx])
                    energy_text = ax.text(center_x, y-text_space-text_shape_space_y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})

                elif  target_location == 's': #cs  
                    if target[idx]:
                        combine_text = f'{original_y[idx]} {target[idx]}'
                    else:
                        combine_text = str(original_y[idx])
                    energy_text = ax.text(center_x, y+text_space+text_shape_space_y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})
                elif  target_location == 'a': #ca
                    if target[idx]:
                        combine_text = f'{original_y[idx]} {target[idx]}'
                    else:
                        combine_text = str(original_y[idx])
                    energy_text = ax.text(center_x-text_space-text_shape_space-len(combine_text)/2*font_size*0.5, y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})

                elif  target_location == 'd': #cd     
                    if target[idx]:
                        combine_text = f'{original_y[idx]} {target[idx]}'
                    else:
                        combine_text = str(original_y[idx])
                    energy_text = ax.text(center_x+text_space+text_shape_space+len(combine_text)/2*font_size*0.5, y , combine_text, color=text_color,fontsize=fontsize, ha='center', va='center_baseline',fontdict={'family': 'Arial'})

    class CustomNavigationToolbar(NavigationToolbar2Tk):
        def __init__(self, canvas, parent):
            super().__init__(canvas, parent)
            self.ax = ax
            self.show_all = show_all
            self.auto_zoom = auto_zoom
            self.zoom_mode_active = False  # 标记是否处于缩放模式
            self._zoom_timer = None  # 定时器ID
            self._is_dragging = False  # 标记是否正在右键拖动

            # 连接 Matplotlib 的 button_release_event
            canvas.mpl_connect("button_release_event", self.on_button_release)
            canvas.mpl_connect("button_press_event", self.on_button_press)

        def home(self, *args):
            """扩展'Home'按钮行为"""
            # 调用父类的home方法，执行原本的Home行为
            super().home(*args)
            
            # 执行自定义的show_all函数
            self.show_all()

        def forward(self, *args):
            """重写'Forward'按钮行为"""
            # 调用父类的forward方法，执行原本的Forward行为
            super().forward(*args)

            # 执行自定义的auto_zoom函数
            self.auto_zoom()

        def back(self, *args):
            """重写'Back'按钮行为"""
            # 调用父类的back方法，执行原本的Back行为
            super().back(*args)

            # 执行自定义的auto_zoom函数
            self.auto_zoom()

        def zoom(self, *args):
            """重写'Zoom to Rectangle'按钮行为"""
            super().zoom(*args)  # 调用默认的放大镜功能
            self.zoom_mode_active = True  # 标记进入缩放模式

        def on_button_press(self, event):
            """鼠标按下事件处理函数"""
            if event.button == 3:  # 右键按下
                self._is_dragging = True  # 标记开始拖动
                self._schedule_zoom()  # 启动定时器，每隔0.1秒执行auto_zoom()

        def on_button_release(self, event):
            self.auto_zoom()
            """鼠标释放事件处理函数"""
            if self._is_dragging:  # 右键释放
                self._is_dragging = False  # 标记停止拖动
                self._cancel_zoom()  # 停止定时器
                self.auto_zoom()  # 释放时调用一次auto_zoom()

        def _schedule_zoom(self):
            """每隔0.1秒调用一次auto_zoom()"""
            if self._zoom_timer is None:
                self._zoom_timer = self.canvas.get_tk_widget().after(10, self._zoom_callback)

        def _zoom_callback(self):
            """定时器回调函数"""
            self.auto_zoom()  # 执行自动缩放
            if self._is_dragging:  # 如果拖动还在进行，继续调用定时器
                self._zoom_timer = self.canvas.get_tk_widget().after(10, self._zoom_callback)

        def _cancel_zoom(self):
            """停止定时器"""
            if self._zoom_timer is not None:
                self.canvas.get_tk_widget().after_cancel(self._zoom_timer)
                self._zoom_timer = None

        def configure_subplots(self):
            """重写配置按钮行为"""
            tool = super().configure_subplots()  # 调用父类的方法，显示默认配置工具

            # 遍历控件，绑定滑块事件
            for attr in dir(tool):
                obj = getattr(tool, attr)
                if isinstance(obj, Slider):  # 检查是否是滑块
                    obj.on_changed(self.on_slider_change)

        def on_slider_change(self, val):
            """滑块改变时触发的回调函数"""
            self.auto_zoom()

        def save_figure(self):
            """重写保存按钮行为"""

            # 弹出对话框，获取用户输入的 DPI 值
            dpi = simpledialog.askinteger("Save Figure", "Enter DPI (e.g., 100, 200, 300):", minvalue=50, maxvalue=1000)
            if dpi is None:
                print("Save canceled by user")
                return

            # 弹出文件保存对话框
            file_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                    filetypes=[("PNG files", "*.png"),
                                                            ("JPEG files", "*.jpg"),
                                                            ("All files", "*.*")])
            if not file_path:
                print("Save canceled by user")
                return

            # 保存图像
            self.canvas.figure.savefig(file_path, dpi=dpi)
            print(f"Figure saved to {file_path} with DPI {dpi}")

    def show_all():
        ax.set_xlim(-1, 1)
        ax.set_ylim(-1, 1)
        update_plot()

    def auto_zoom():
        font_size = font_size_var.get()
        curve_width = curve_width_var.get()
        bond_width = bond_width_var.get()

        x_min, x_max = ax.get_xlim()
        y_min, y_max = ax.get_ylim()

        # 获得X轴方向和Y轴方向的像素数
        width, height = canvas.get_width_height()

        # 缩放因子计算
        x_range = x_max - x_min
        y_range = y_min - y_max
        
        # 控制页面物理尺寸的精华部分
        fontsize = font_size/max(x_range/width, y_range/height)*0.7
        hline_linewidth = bond_width/max(x_range/width, y_range/height)*0.7
        curve_linewidth = curve_width/max(x_range/width, y_range/height)*0.7
        size_marker_linewidth = 0.6/max(x_range/width, y_range/height)*0.7


        for text_obj in ax.texts:
            # current_fontsize = text_obj.get_fontsize()
            # print(current_fontsize)
            text_obj.set_fontsize(fontsize)

        for hline_obj in hlines:
            hline_obj.set_linewidth(hline_linewidth)

        for curve_obj in curve_lines:
            for curve in curve_obj:
                curve.set_linewidth(curve_linewidth)

        for size_marker_obj in size_markers:
            size_marker_obj.set_linewidth(size_marker_linewidth)

        # 重新绘制画布
        canvas.draw()

    def on_resize(event, ax, canvas):
        """在窗口大小变化时触发"""
        auto_zoom()

    def toggle_grid():
        # 切换网格显示
        ax.grid(grid_var.get())
        canvas.draw()
        
    # Create main window
    root = tk.Tk()
    root.title("Interactive Bezier Curve Adjustments")
    style = ttk.Style(root)
    style.theme_use("vista")  # 使用 'vista' 主题
    main_frame = tk.Frame(root)
    main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)



    # Add Matplotlib figure
    global fig
    fig = Figure(figsize=(8, 6),dpi=100)
    ax = fig.add_subplot(111)
    # ax.set_title("Interactive Bezier Curve")
    # ax.set_xlabel("Adjusted Coordinate")
    # ax.set_ylabel("Potential Energy")
    ax.tick_params(axis='both', which='both', bottom=False, top=False,
               left=False, right=False, labelbottom=False, labelleft=False)
    ax.grid(True)
    canvas = FigureCanvasTkAgg(fig, master=main_frame)
    toolbar = CustomNavigationToolbar(canvas, main_frame)

    toolbar.update()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    canvas.mpl_connect("resize_event", lambda event: on_resize(event, ax, canvas))
    fig.tight_layout() 

    
    def on_zoom(event):
        """
        根据缩放动态调整文本(ax.text)的字体大小，同时调整坐标范围。
        支持放大和缩小
        """
        x_min, x_max = ax.get_xlim()
        y_min, y_max = ax.get_ylim()
        font_size = font_size_var.get()
        curve_width = curve_width_var.get()
        bond_width = bond_width_var.get()

        # 缩放因子计算
        x_range = x_max - x_min
        y_range = y_max - y_min

        # 根据鼠标滚轮事件确定缩放方向
        if event.button == 'up':  # 放大
            x_min_new = x_min + x_range * 0.1
            x_max_new = x_max - x_range * 0.1
            y_min_new = y_min + y_range * 0.1
            y_max_new = y_max - y_range * 0.1
        elif event.button == 'down':  # 缩小
            x_min_new = x_min - x_range * 0.1
            x_max_new = x_max + x_range * 0.1
            y_min_new = y_min - y_range * 0.1
            y_max_new = y_max + y_range * 0.1

        else:
            return  # 非缩放事件直接返回
        
        # 更新坐标范围
        ax.set_xlim(x_min_new, x_max_new)
        ax.set_ylim(y_min_new, y_max_new)

        x_min, x_max = ax.get_xlim()
        y_min, y_max = ax.get_ylim()

        # 获得X轴方向和Y轴方向的像素数
        width, height = canvas.get_width_height()

        # 缩放因子计算
        x_range = x_max - x_min
        y_range = y_min - y_max
        
        # 控制页面物理尺寸的精华部分
        fontsize = font_size/max(x_range/width, y_range/height)*0.7
        hline_linewidth = bond_width/max(x_range/width, y_range/height)*0.7
        curve_linewidth = curve_width/max(x_range/width, y_range/height)*0.7
        size_marker_linewidth = 0.6/max(x_range/width, y_range/height)*0.7


        for text_obj in ax.texts:
            # current_fontsize = text_obj.get_fontsize()
            # print(current_fontsize)
            text_obj.set_fontsize(fontsize)

        for hline_obj in hlines:
            hline_obj.set_linewidth(hline_linewidth)

        for curve_obj in curve_lines:
            for curve in curve_obj:
                curve.set_linewidth(curve_linewidth)

        for size_marker_obj in size_markers:
            size_marker_obj.set_linewidth(size_marker_linewidth)

        # 重新绘制画布
        canvas.draw()
    
    class RightClickMenu:
        def __init__(self, table_frame, table):
            self.table_frame = table_frame
            self.table = table
            # Bind right-click event
            self.columns = self.table["columns"]
            self.table.bind("<Button-3>", self.show_context_menu)

            # Create context menu
            self.menu = tk.Menu(root, tearoff=0)
            self.menu.add_command(label="Delete Row", command=self.delete_row)
            self.menu.add_command(label="Delete Column", command=self.delete_column)
            self.menu.add_command(label="Add Row (above)", command=self.add_row)
            self.menu.add_command(label="Add Column (left)", command=self.add_column)

            self.selected_item = None
            self.selected_column = None

        def show_context_menu(self, event):
            # Get the row and column at click position
            region = self.table.identify("region", event.x, event.y)
            if region == "cell":
                row_id = self.table.identify_row(event.y)
                column = self.table.identify_column(event.x)

                if row_id and column:
                    col_index = int(column.replace("#", ""))
                    if col_index >= 4:
                        self.selected_item = row_id
                        self.selected_column = column
                        # Highlight the selected row
                        self.table.selection_set(row_id)
                        # Show context menu
                        self.menu.post(event.x_root, event.y_root)

        def delete_row(self):
            if self.selected_item:
                self.table.delete(self.selected_item)
                self.selected_item = None

        def delete_column(self):
            if self.selected_column:
                existing_columns = list(self.table['columns'])
                if len(existing_columns) <= 3:
                    print("无法删除：至少需要保留3列。")
                    return
                col_index = int(self.selected_column.replace("#", "")) - 1
                existing_columns.pop(col_index)
                self.columns = existing_columns[:3] + [f"E{i-3}" for i in range(4, len(existing_columns) + 1)]
                self.table['columns'] = tuple(self.columns)
                default_column_width = 100
                for i, col in enumerate(self.columns):
                    self.table.heading(col, text=col)
                    self.table.column(col, width=default_column_width, anchor='center', stretch=False)
                for item in self.table.get_children():
                    values = list(self.table.item(item, 'values'))
                    if len(values) > 0:
                        del values[col_index]
                    self.table.item(item, values=values)

        def add_row(self):
            if self.selected_item:
                index = self.table.index(self.selected_item)
                num_columns = len(self.table['columns'])
                new_row_values = ('#000000', '#000000', '#000000') + ('',) * (num_columns - 3)
                self.table.insert('', index, values=new_row_values)

        def add_column(self):
            existing_columns = list(self.table['columns'])
            new_column = f'E{len(existing_columns) - 3}'
            if self.selected_column:
                col_index = int(self.selected_column.replace("#", "")) - 1
                existing_columns.insert(col_index, new_column)
                self.columns = existing_columns[:3] + [f"E{i-3}" for i in range(4, len(existing_columns) + 1)]
                self.table['columns'] = tuple(self.columns)
                default_column_width = 100
                for col in self.columns:
                    self.table.heading(col, text=col)
                    self.table.column(col, width=default_column_width, anchor='center', stretch=False)
                for item in self.table.get_children():
                    values = list(self.table.item(item, 'values'))
                    values.insert(col_index, '')
                    self.table.item(item, values=values)

    # 绑定放大和缩小事件处理函数到 Matplotlib 图表上
    fig.canvas.mpl_connect('scroll_event', on_zoom)

    def create_table_window():
        # Create a separate window for the table
        table_window = tk.Toplevel(root)
        table_window.resizable(True, True) 
        table_window.title("Energy Data Table")

        # Create the sheet
        sheet = tksheet.Sheet(table_window)
        sheet.enable_bindings((
            "single_select",
            "row_select",
            "column_select",
            "drag_select",
            "select_all",
            "column_width_resize",
            "arrowkeys",
            "right_click_popup_menu",
            "rc_select",
            "rc_insert_row",
            "rc_delete_row",
            "copy",
            "cut",
            "paste",
            "delete",
            "undo",
            "edit_cell"
        ))

        # Set initial data
        initial_data = [['#000000', '#000000', '#000000', '0.0', '0.0', '0.0']]
        headers = ['Curve Color', 'Marker Color', 'Text Color', 'E1', 'E2', 'E3']
        sheet.headers(headers)
        sheet.data = initial_data

        # Set column widths
        for i in range(len(headers)):
            sheet.column_width(column=i, width=100)

        sheet.pack(fill="both", expand=True, padx=5, pady=5)

        # Right click menu functionality
        def right_click_menu(event):
            popup = tk.Menu(table_window, tearoff=0)
            popup.add_command(label="Add Row", command=add_row)
            popup.add_command(label="Add Column", command=add_column)
            popup.add_command(label="Delete Row", command=delete_row)
            popup.add_command(label="Delete Column", command=delete_column)
            popup.add_separator()
            popup.add_command(label="Save Data", command=save_table_data)
            popup.add_command(label="Load Data", command=load_table_data)

            try:
                popup.tk_popup(event.x_root, event.y_root, 0)
            finally:
                popup.grab_release()

        sheet.bind("<Button-3>", right_click_menu)

        # Handle cell editing for color columns
        def on_cell_edit(event):
            r, c = event.row, event.column
            current_value = sheet.get_cell_data(r, c)

            if c in [0, 1, 2]:  # Color columns
                if current_value.startswith('#'):
                    default_color = current_value
                else:
                    default_color = "#000000"

                color_code = colorchooser.askcolor(initialcolor=default_color, title="选择颜色")[1]
                if color_code:
                    sheet.set_cell_data(r, c, value=color_code)
                    update_plot()
                return "break"  # Prevent default text editing

            elif c > 2:  # Numeric columns
                if current_value == '' or current_value == '0.0':
                    show_all()

        sheet.bind("<<SheetModified>>", on_cell_edit)

        # Add column function
        def add_column():
            current_headers = sheet.headers()
            new_col_name = f'E{len(current_headers) - 2}'
            new_headers = current_headers + [new_col_name]
            sheet.headers(new_headers)

            # Add empty data for new column
            new_data = []
            for row in sheet.get_sheet_data():
                new_row = list(row) + ['']
                new_data.append(new_row)
            sheet.data = new_data

            # Set column width
            sheet.column_width(column=len(new_headers)-1, width=100)

        # Add row function
        def add_row():
            num_cols = len(sheet.headers())
            new_row = ['#000000', '#000000', '#000000'] + [''] * (num_cols - 3)
            sheet.insert_row(values=new_row)

        # Delete column function
        def delete_column():
            current_headers = sheet.headers()
            if len(current_headers) <= 3:
                print("无法删除：至少需要保留3列。")
                return

            # Get data without last column
            new_data = []
            for row in sheet.get_sheet_data():
                new_row = list(row)[:-1]
                new_data.append(new_row)

            # Update headers and data
            sheet.headers(current_headers[:-1])
            sheet.data = new_data

        # Delete row function
        def delete_row():
            if len(sheet.get_sheet_data()) > 0:
                sheet.delete_row()

        # Save table data function
        def save_table_data():
            if os.path.exists('table_data.json'):
                response = messagebox.askyesno("Confirm", "The file 'table_data.json' already exists. Do you want to overwrite it?")
                if not response:
                    print("Save operation cancelled.")
                    return

            data = sheet.get_sheet_data()
            with open('table_data.json', 'w') as file:
                json.dump(data, file)
            print(f"Table data saved at {os.getcwd()}\\table_data.json.")

        # Load table data function
        def load_table_data():
            try:
                with open('table_data.json', 'r') as file:
                    data = json.load(file)

                if data:
                    # Dynamically set headers
                    headers = ['Curve Color', 'Marker Color', 'Text Color'] + [f'E{i+1}' for i in range(len(data[0])-3)]
                    sheet.headers(headers)
                    sheet.data = data
                    show_all()

                print(f"Table data loaded from {os.getcwd()}\\table_data.json.")
            except FileNotFoundError:
                print(f"No saved data found at {os.getcwd()}.")

        # Create buttons frame
        buttons_frame = tk.Frame(table_window)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)

        # Create buttons
        add_row_button = tk.Button(buttons_frame, text="Add Row", command=add_row)
        add_column_button = tk.Button(buttons_frame, text="Add Column", command=add_column)
        delete_row_button = tk.Button(buttons_frame, text="Delete Row", command=delete_row)
        delete_column_button = tk.Button(buttons_frame, text="Delete Column", command=delete_column)
        save_button = tk.Button(buttons_frame, text="Save Data", command=save_table_data)
        load_button = tk.Button(buttons_frame, text="Load Data", command=load_table_data)

        # Pack buttons
        add_row_button.pack(side=tk.LEFT, padx=5, pady=5)
        add_column_button.pack(side=tk.LEFT, padx=5, pady=5)
        delete_row_button.pack(side=tk.LEFT, padx=5, pady=5)
        delete_column_button.pack(side=tk.LEFT, padx=5, pady=5)
        save_button.pack(side=tk.LEFT, padx=5, pady=5)
        load_button.pack(side=tk.LEFT, padx=5, pady=5)

        return sheet

    global table

    table = create_table_window()
    # 存储所有动态创建的滑条
    slider_frames = {}

    def add_slider_with_entry(label, variable, from_, to, resolution, row, length=100):

        #检查是否已有相同 label 的滑条，若存在则销毁
        if label in slider_frames:
            slider_frames[label].destroy()
            del slider_frames[label]

        # 创建一个新的 Frame 用于包装每一组 slider 和 entry
        row_frame = tk.Frame(control_frame,height=10)
        row_frame.pack(fill='y', padx=5, pady=5,expand=True)

        #存储滑条引用
        slider_frames[label] = row_frame 
        # 设置标签宽度，确保它们具有一致的宽度
        label_widget = ttk.Label(row_frame, text=label, width=12)  # 设定固定宽度
        label_widget.pack(side=tk.LEFT, padx=5, anchor='s')

        # 确定分辨率的小数位数
        precision = len(str(resolution).split('.')[-1]) if '.' in str(resolution) else 0

        # 创建一个减少按钮
        def decrease_value():
            new_value = round(variable.get() - resolution, precision)
            if new_value >= from_:
                variable.set(new_value)

        def decrease_hold():
            decrease_value()
            global hold_decrease
            hold_decrease = row_frame.after(50, decrease_hold)  # 50ms 后继续调用

        def stop_decrease():
            global hold_decrease
            row_frame.after_cancel(hold_decrease)

        decrease_button = ttk.Button(row_frame, text='-', width=2)
        decrease_button.pack(side=tk.LEFT, padx=2, anchor='s')
        decrease_button.bind("<ButtonPress-1>", lambda e: decrease_hold())
        decrease_button.bind("<ButtonRelease-1>", lambda e: stop_decrease())

        # 创建滑动条并添加到内部 Frame 使用 pack 布局
        slider = tk.Scale(row_frame, from_=from_, to=to, variable=variable, orient=tk.HORIZONTAL,
                        resolution=resolution, length=length, showvalue=False)
        slider.pack(side=tk.LEFT, padx=5, pady=2, anchor='s')  # 使用 pack 来放置滑动条

        # 创建一个增加按钮
        def increase_value():
            new_value = round(variable.get() + resolution, precision)
            if new_value <= to:
                variable.set(new_value)

        def increase_hold():
            increase_value()
            global hold_increase
            hold_increase = row_frame.after(50, increase_hold)  # 50ms 后继续调用

        def stop_increase():
            global hold_increase
            row_frame.after_cancel(hold_increase)

        increase_button = ttk.Button(row_frame, text='+', width=2)
        increase_button.pack(side=tk.LEFT, padx=2, anchor='s')
        increase_button.bind("<ButtonPress-1>", lambda e: increase_hold())
        increase_button.bind("<ButtonRelease-1>", lambda e: stop_increase())

        # 创建输入框并添加到当前行
        entry = ttk.Entry(row_frame, textvariable=variable, width=8)
        entry.pack(side=tk.LEFT, padx=5, anchor='s')  # 保持输入框使用 pack 布局

        # 为输入框绑定事件，按回车键时更新图形
        entry.bind("<Return>", update_plot)

        # 为滑块和输入框添加事件，使得值变化时都能更新图形
        variable.trace_add("write", update_plot)
        
        
    # 创建主窗口和控制面板
    control_frame = tk.Frame(root)
    control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

    # 配置控件所在列的权重
    control_frame.columnconfigure(1, weight=1)

    # 定义可调节的变量
    adjustment_factor_var = tk.DoubleVar(value=0.0)
    width_factor_var = tk.DoubleVar(value=0.3)
    adjustment_factor_2_var = tk.DoubleVar(value=0.0)
    scale_factor_var = tk.DoubleVar(value=10.0)
    scale_factor_x_var = tk.DoubleVar(value=1.0)
    scale_factor_y_var = tk.DoubleVar(value=1.0)
    bond_length_var = tk.DoubleVar(value=20.0)
    radius_var = tk.DoubleVar(value=3.0)
    font_size_var = tk.DoubleVar(value=10.0)
    text_space_var = tk.DoubleVar(value=3.0)
    curve_width_var = tk.DoubleVar(value=0.6)
    bond_width_var = tk.DoubleVar(value=2.0)

    # 添加滑块和输入框
    add_slider_with_entry("Scale Factor", scale_factor_var, 0, 20.0, 0.1, 3)
    add_slider_with_entry("Scale Factor x", scale_factor_x_var, 0.0, 5, 0.02, 3)
    add_slider_with_entry("Scale Factor y", scale_factor_y_var, 0.0, 10.0, 0.1, 3)
    add_slider_with_entry("Base Curve", width_factor_var, 0.0, 1.0, 0.01, 0)
    add_slider_with_entry("Disp Factor", adjustment_factor_var, 0.0, 2, 0.01, 1)
    add_slider_with_entry("Curve Factor", adjustment_factor_2_var, 0.0, 0.02, 0.0001, 2)
    add_slider_with_entry("Curve Width", curve_width_var, 0.0, 5.0, 0.1, 2)
    add_slider_with_entry("Font Size", font_size_var, 0.0, 100.0, 1, 6)
    add_slider_with_entry("Text Space", text_space_var, 0.0, 50, 1, 6)

    spacer = tk.Frame(control_frame, height=10)  # 创建一个空白 Frame，高度为 10
    spacer.pack()

    # 创建形状选择框
    shape_frame = tk.Frame(control_frame)
    shape_frame.pack(padx=5, pady=2, anchor='w',fill='y',expand=True)

    tk.Label(shape_frame, text="Shape:").pack(side=tk.LEFT, padx=5)
    shape_type_var = tk.StringVar(value="line")
    tk.Radiobutton(shape_frame, text="Line", variable=shape_type_var, value="line", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(shape_frame, text="Circle", variable=shape_type_var, value="circle", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(shape_frame, text="None", variable=shape_type_var, value="None", command=update_plot).pack(side=tk.LEFT, padx=10)

    add_slider_with_entry("Bond Length", bond_length_var, 0.0, 40.0, 0.5, 4)
    add_slider_with_entry("Bond Width", bond_width_var, 0.0, 10.0, 0.1, 2)
    add_slider_with_entry("Circle Radius", radius_var, 0.0, 10.0, 0.1, 5)

    # 创建连接选择框
    connection_frame = tk.Frame(control_frame)
    connection_frame.pack(padx=5, pady=5, anchor='w',fill='y',expand=True)

    tk.Label(connection_frame, text="Connection:").pack(side=tk.LEFT, padx=5)
    connect_type_var = tk.StringVar(value="center")
    tk.Radiobutton(connection_frame, text="Center", variable=connect_type_var, value="center", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(connection_frame, text="Side", variable=connect_type_var, value="side", command=update_plot).pack(side=tk.LEFT, padx=5)

    # 创建线性选择框
    line_type_frame = tk.Frame(control_frame)
    line_type_frame.pack(padx=5, pady=5, anchor='w',fill='y',expand=True)

    tk.Label(line_type_frame, text="Line Type:").pack(side=tk.LEFT, padx=5)
    line_type_var = tk.StringVar(value="Solid")
    tk.Radiobutton(line_type_frame, text="Solid", variable=line_type_var, value="Solid", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(line_type_frame, text="Dashed", variable=line_type_var, value="Dashed", command=update_plot).pack(side=tk.LEFT, padx=5)

    # 创建能量和标签的连接形式框
    target_layout_frame = tk.Frame(control_frame)
    target_layout_frame .pack(padx=5, pady=5, anchor='w',fill='y',expand=True)

    tk.Label(target_layout_frame , text="Label-Energy Layout").pack(side=tk.LEFT, padx=5)
    target_layout_var = tk.StringVar(value="sperate")
    tk.Radiobutton(target_layout_frame, text="Sperate", variable= target_layout_var, value="sperate", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(target_layout_frame, text="Combine", variable= target_layout_var, value="combine", command=update_plot).pack(side=tk.LEFT, padx=5)

    # 创建标签位置选择框
    target_location_frame = tk.Frame(control_frame)
    target_location_frame .pack(padx=5, pady=5, anchor='w',fill='y',expand=True)

    tk.Label(target_location_frame , text="Target Location").pack(side=tk.LEFT, padx=5)
    target_location_var = tk.StringVar(value="c")
    tk.Radiobutton(target_location_frame, text="C", variable= target_location_var, value="c", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(target_location_frame, text="W", variable= target_location_var, value="w", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(target_location_frame, text="S", variable= target_location_var, value="s", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(target_location_frame, text="A", variable= target_location_var, value="a", command=update_plot).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(target_location_frame, text="D", variable= target_location_var, value="d", command=update_plot).pack(side=tk.LEFT, padx=5)

    # 使用一个子Frame排列选项框
    options_frame = tk.Frame(control_frame)
    options_frame.pack(fill=tk.X, expand=True)

    show_marker_var = tk.BooleanVar(value=False)
    checkbutton = tk.Checkbutton(options_frame, text="Show Marker", variable=show_marker_var, command=update_plot)
    checkbutton.pack(side=tk.LEFT,padx=5, pady=5, anchor='w',fill='y')

    grid_var = tk.BooleanVar(value=True)
    grid_checkbutton = tk.Checkbutton(options_frame, text="Show Grid", variable=grid_var, command=update_plot)
    grid_checkbutton.pack(side=tk.LEFT,padx=5, pady=5, anchor='w',fill='y')


    def hex_to_rgb(hex_color):
        """将16进制颜色代码转换为RGB元组"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) / 255.0 for i in (0, 2, 4))

    def add_colors_to_colortable(color_list):
        """将颜色列表添加到colortable中，保留原色"""
        # 创建根元素 <colortable>
        colortable = ET.Element("colortable")
        
        # 保留原色
        original_colors = [
            "#FFFFFF",  # white
            "#000000",  # black
            "#FF0000",  # red
            "#FFFF00",  # yellow
            "#00FF00",  # green
            "#00FFFF",  # cyan
            "#0000FF",  # blue
            "#FF00FF"   # magenta
        ]
        
        # 添加原色
        for color in original_colors:
            r, g, b = hex_to_rgb(color)
            color_element = ET.SubElement(colortable, "color", r=str(r), g=str(g), b=str(b))
        
        # 添加额外的颜色列表中的颜色
        for color in color_list:
            r, g, b = hex_to_rgb(color)
            color_element = ET.SubElement(colortable, "color", r=str(r), g=str(g), b=str(b))
        
        # 创建一个ElementTree对象并生成XML字符串
        xml_str = ET.tostring(colortable, encoding="unicode", method="xml")
        
        return xml_str
    
    def export_cdxml(*args):
        global already_draw_target
        already_draw_target = []
        try:
            # Get user input parameters
            adjustment_factor = float(adjustment_factor_var.get())
            width_factor = float(width_factor_var.get())
            adjustment_factor_2 = float(adjustment_factor_2_var.get())
            scale_factor = float(scale_factor_var.get())
            scale_factor_x = float(scale_factor_x_var.get())
            scale_factor_y = float(scale_factor_y_var.get())
            bond_length = float(bond_length_var.get())
            radius = float(radius_var.get())
            connect_type = connect_type_var.get()
            shape_type = shape_type_var.get()
            line_type = line_type_var.get()
            font_size = font_size_var.get()
            text_space = text_space_var.get()
            target_layout = target_layout_var.get()
            target_location = target_location_var.get()
            curve_width = curve_width_var.get()
            bond_width = bond_width_var.get()
            

            connect_xml = ''
            graph_xml = ''
            color_list  = [] # color target
            global page_high
            global page_width
            page_high = 1
            page_width = 1
            y_limit = 300

            for row in table.get_children():
                values = table.item(row)['values']
                for i in values[:3]: # color target
                    color_list.append(i)

            color_list = list(set(color_list)) # remove same color
            color_xml = add_colors_to_colortable(color_list)

            # set y limit
            total_y= []
            for row in table.get_children():
                values = table.item(row)['values']
                for index, value in enumerate(values[3:]):
                    value = str(value)
                    if value.strip():  # If value is not empty
                        try:
                            value = parse_data(value)
                            total_y.append(float(value[0]))

                        except ValueError:
                            print(f"Invalid value at row {row}, column {index + 2}. Please correct it.")
                            return
                        
            scaled_y_points = [y * scale_factor*scale_factor_y / 2 for y in total_y]

            for i in scaled_y_points:
                if i > y_limit :
                    y_limit = i+100

            # generate curve and target data
            for row in table.get_children():
                values = table.item(row)['values']
                original_y = []
                target = []
                location = []
                original_x = []
                for index, value in enumerate(values[3:]):
                    value = str(value)
                    if value.strip():  # If value is not empty
                        try:
                            value = parse_data(value)
                            original_y.append(float(value[0]))
                            target.append(value[1])
                            location.append(value[2])
                            original_x.append(index + 1)
                            
                        except ValueError:
                            print(f"Invalid value at row {row}, column {index + 2}. Please correct it.")
                            return
                        
                curve_color = values[0]
                marker_color = values[1]
                text_color = values[2]

                
                # Call Bezier curve calculation function (dummy function here)
                x_adjusted, y_points, bezier_curves = get_bezier_curve_points_flat(
                   original_x, original_y, adjustment_factor, width_factor, adjustment_factor_2
                )
                
                # Scale Bezier curve
                if  x_adjusted[0] == 'no_data': # pass empty line
                    continue
                # Scale Bezier curve
                for curve in bezier_curves:
                    X = curve['X']
                    Y = curve['Y']
                    for i in range(len(X)):
                        X[i] = X[i] * scale_factor *scale_factor_x* 10 + 50
                        Y[i] = y_limit - Y[i] * scale_factor*scale_factor_y / 2

                # Scale adjusted coordinates and potential energy values
                scaled_x_adjusted = [x * scale_factor*scale_factor_x * 10 + 50 for x in x_adjusted]
                scaled_y_points = [y_limit - y * scale_factor*scale_factor_y / 2 for y in y_points]

                # auto change page width and high
                for x in  scaled_x_adjusted:
                    ad_page_width = math.ceil(abs(x/510))
                    if ad_page_width > page_width:
                        page_width = ad_page_width
                for y in scaled_y_points:
                    ad_page_high = math.ceil(abs(y/710))
                    if ad_page_high > page_high:
                        page_high = ad_page_high

                curve_color_index = color_list.index(curve_color) + 10
                marker_color_index = color_list.index(marker_color) + 10
                text_color_index = color_list.index(text_color) + 10

                text_base_movement = round(font_size*0.3241+0.1236,2)

                
                if shape_type ==  "line":
                    if width_factor == 0:
                        curves_xml = draw_line(scaled_x_adjusted,scaled_y_points,line_type,curve_width,bond_length,curve_color_index,connect_type,original_x)
                    else:
                        curves_xml = draw_curve(bezier_curves, connect_type, bond_length,curve_color_index,line_type,original_x,curve_width)
                    graph_xml  += "\n".join(curves_xml)
                    rectangle_xml = generate_rectangle_xml(scaled_x_adjusted, scaled_y_points, bond_length, marker_color_index, connect_type,original_x,bond_width)
                    graph_xml  += "\n".join(rectangle_xml)
                    text_xml = generate_text_cdxml(scaled_x_adjusted, scaled_y_points, original_y,target, location,target_layout,target_location,text_space,text_base_movement,bond_length,bond_width/2,text_color_index, connect_type, original_x,font_size,font_type=3, Z_value=70)
                    graph_xml  += "\n".join(text_xml)

                elif shape_type == "circle":
                    if width_factor == 0:
                        curves_xml = draw_line(scaled_x_adjusted,scaled_y_points,line_type,curve_width,radius,curve_color_index,connect_type,original_x)
                    else:
                        curves_xml = draw_curve(bezier_curves, connect_type, radius,curve_color_index,line_type,original_x,curve_width)
                    graph_xml  += "\n".join(curves_xml)
                    circles_xml = generate_circle_xml(scaled_x_adjusted, scaled_y_points, radius, marker_color_index, connect_type,original_x)
                    graph_xml  += "\n".join(circles_xml)
                    text_xml = generate_text_cdxml(scaled_x_adjusted, scaled_y_points, original_y,target, location,target_layout,target_location,text_space,text_base_movement,radius,radius,text_color_index, connect_type, original_x,font_size,font_type=3, Z_value=70)
                    graph_xml  += "\n".join(text_xml)
                elif shape_type == "None":
                    if width_factor == 0:
                        curves_xml = draw_line(scaled_x_adjusted,scaled_y_points,line_type,curve_width,0,curve_color_index,connect_type,original_x)
                    else:
                        curves_xml = draw_curve(bezier_curves, connect_type,0,curve_color_index,line_type,original_x,curve_width)
                    graph_xml  += "\n".join(curves_xml)
                    text_xml = generate_text_cdxml(scaled_x_adjusted, scaled_y_points, original_y,target, location,target_layout,target_location,text_space,text_base_movement,0,0,text_color_index, connect_type, original_x,font_size,font_type=3, Z_value=70)
                    graph_xml  += "\n".join(text_xml)

            page_xml = f'''<page
                id="66"
                BoundingBox="0 0 1620 719.75"
                HeaderPosition="36"
                FooterPosition="36"
                PrintTrimMarks="yes"
                HeightPages="{page_high}"
                WidthPages="{page_width}"
                >'''

            connect_xml  += "".join(color_xml)  # add color
            connect_xml  += "".join(font_xml)  # add font
            connect_xml  += "".join(page_xml) # add page
            connect_xml  += "".join(graph_xml) # add graph

            # Combine all parts into a single CDXML string
            cdxml_string = cdxml_header + connect_xml + cdxml_footer

            # Save CDXML file
            save_cdxml_file(cdxml_string)
            print(f"The CDXML file has been successfully generated: {os.getcwd()}\\output.cdxml.")
            
        except ValueError:
            print("Input Error", "Please enter valid numeric values.")
    
    def re_create_table_window():
        global table
        # 检查 table 是否已存在且未关闭
        if table is None or not table.winfo_exists():
            # 如果没有窗口或者窗口已经被关闭，创建新的窗口
            table = create_table_window()
        else:
            print("Table window is already open.")
    export_button = tk.Button(control_frame, text="Show All", command=show_all)
    export_button.pack(side=tk.LEFT, padx=5, pady=20)

    export_button = tk.Button(control_frame, text="Export CDXML", command=export_cdxml)
    export_button.pack(side=tk.RIGHT, padx=20, pady=5)

    export_button = tk.Button(control_frame, text="Open Table", command=re_create_table_window)
    export_button.pack(side=tk.LEFT, padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
    interactive_bezier_curve()


