# -*- coding: utf-8 -*-
"""
Created on Mon Mar 28 14:17:52 2022

Trying to script a way to build .twf files from .hdr files produced by the
Saint Mary's EDS scanner.


Okay... so I did that. The new goal of this script is to batch process moving 
all of the sem data from single repositories to qgis projects. 

Update 2022/04/19 - I can batch process most of the SEM information. There are
some repository requirements. I.e. there can only be one .xlsx file. There also
can't be existing tfw files.

DEV NOTE:
    self.organize_sheet is not functioning right now. I believe calling of the 
    spreadsheet was previously removed for some reason. I will look to fix this
    at some point in the future, but for now just ignore the .xlsx part of this
    script.
"""


class mica():
    def __init__(self, path = None, output_path = None):
        self.path = path
        self.output_path = output_path

        
    def find_files(self, file_suffix):
        '''
        Small supporting method, meant to find all files of a type

        Parameters
        ----------
        file_suffix : string, e.g., .tif, .svg
            The filetype to be found

        Returns
        -------
        output : list
            A list of all files in the repository pointed to by "path" that share
            the file type entered with "file_suffix"
        '''
        import glob       
        output = glob.glob(self.path+'/*'+file_suffix)
        return output

    def to_tfw(self, hdr = 'blank'):
        '''
        Reads a -tiff.hdr file, which is produced by the Saint Mary's University
        SEM when snapping photomicrographs, and writes a .tfw file with the 
        pertinent information. For note, a .tfw file contains spacial information 
        for an associated .tif file. The spatial information is, more specifically,
        six lines of text containing the following
        
        The pixel scale of the x direction
        Rotation about the y-axis
        Rotation about the x-axis
        Pixel size in the y-direction
        X-coordinate of the center of the upper left pixel
        Y-coordinate of the upper left pixel
        
        Parameters
        ----------
        path : str
        Filename of the .tif-hdr file.
        '''
        x_scale = 0
        y_scale = 0
        out = hdr.replace('-tif.hdr', '.tfw')
        if self.output_path != None:
            out = out.replace(self.path, self.output_path)
        with open(out, 'a') as file:
            for line in open(hdr):
                if 'PixelSizeX=' in line:
                    x = line.replace('PixelSizeX=', '')
                    x = x.replace('\n', '')
                    x=float(x)
                    x_scale = x
                    file.write(str(x)+'''
''')
            file.write('''0
''')
            file.write('''0
''')
            for line in open(hdr):
                if 'PixelSizeY=' in line:
                    x = line.replace('PixelSizeY=', '')
                    x = x.replace('\n', '')
                    x=float(x)
                    y_scale = x
                    x=x*(-1)
                    file.write(str(x)+'''
''')            
                elif 'StageX=' in line:
                    x = line.replace('StageX=', '')
                    x = x.replace('\n', '')
                    x=float(x)
                    x=x*(-1)
                    x=x-(x_scale*(640))
                    file.write(str(x)+'''
''')
                elif 'StageY=' in line:
                    x = line.replace('StageY=', '')
                    x = x.replace('\n', '')
                    x=float(x)
                    x=x*(-1)
                    x=x+(y_scale*(480))
                    file.write(str(x)+'''
''')
    
    def batch_tfw(self):
        """
        Translates all .tif-hdr files in the input repository then writes .twf
        files into the output repository. See to_tfw for more information
        """
        x = self.find_files(file_suffix='-tif.hdr')
        for name in x:
            self.to_tfw(hdr = name)
            
    def drop_bar(self, raster):
        """
        Copies an electromicrograph from the input repository to the output 
        repository and simultaneously removes the information bar from the 
        electromicrograph

        Parameters
        ----------
        raster : str
            Filename of the electromicrograph to be moved and edited.
        """
        from PIL import Image
        import numpy as np
        
        im = np.array(Image.open(raster))
        new = np.delete(im, np.s_[960:], 0)
        new_img = Image.fromarray(new)
        name = raster.replace(self.path+"/",'')
        if self.output_path != None:
            new_img.save(self.output_path+"/"+name)
        else:
            print('please set an output path')

    def batch_bar(self):
        """
        Copies all electromicrographs from the input repository to the output 
        repository and simultaneously removes the information bar from the 
        electromicrograph
        """
        x = self.find_files(file_suffix='.tif')
        for name in x:
            name = name.replace('\\','/')
            self.drop_bar(name)

    def move_tif(self, raster):
        """
        Copies an electromicrograph from one repository to another

        Parameters
        ----------
        raster : str
            Filename of the electromicrograph to be moved.
        """
        from PIL import Image
        im = Image.open(raster)
        name = raster.replace(self.path+'/', '')
        if self.output_path != None:
            im.save(self.output_path+"/"+name)
        else:
            print('please set an output path')

    def batch_move(self):
        """
        Copies all electromicrographs from the input repository to the output 
        repository.
        """
        x = self.find_files(file_suffix='.tif')
        for name in x:
            name = name.replace('\\','/')
            self.move_tif(name)
                
    def organize_sheet(self, frame):
        '''
        This function will read data from .xlsx spreadsheets produced by users
        of the SMU SEM. The spreadhseets are discontinuous, made up of sub-sheets
        that can be pseudo-randomly spread throughout the spreadsheet, and of
        varying sizes. As long as each sub-sheet was placed in the first column
        of the spreadsheet, this method will organize the elemental data into
        a single, continuous spreadsheet, without blank spaces. 
        
        No transformations happen to the data during this method, the data is 
        simply organized into a functional spreadsheet
        
        Parameters
        ----------
        excel_file : string
            Full filepath of the excel file containing the spreadsheet that 
            needs to be organized. PLEASE NOTE this must be a .xlsx file.
        sheet : string=Sheet1
            Name of the sheet within the .xlsx file that needs to be organized. 
            The default is 'Sheet1', which should apply to most folks using 
            this on data collected on the SMU SEM.

        Returns
        -------
        output : pandas DataFrame
            A DataFrame containing all of the elemental abundance chemistry from
            the spreadsheet. The index of the dataframe is from 0 to n, n being
            the total number of spectra in the dataset. The columns of the 
            DataFrame are as follows:
            
            sample:   The "ID:" of the data
            site:     The "Site of Interest" number
            Spectrum: The spectrum number
            * Any other columns, including elements measured and X Y coordinates
            in the order they appear in in the sheet
            Total:    The "Total" (?)
        '''
        
        import pandas as pd
        
        # importing the spreadsheet as mess, because its a mess. Also NaN makes
        # everything worse, so we're filling those right up
        mess = frame #import the big messy sheet
        mess = mess.fillna('0')
        
        # Separating out column 1, where the starting and ending key cells occur
        col_1 = mess['Unnamed: 0']
        
        # Creating attributes with a method scale scope (?) to accept and keep
        # information from the if loop
        starts = []
        ends = []
        
        # The if loop identifies staring and ending key cells, which will contain
        # 'Project:' and 'Min.' respectively. These identify sub-sheets to extract
        for line in col_1.index:
            if isinstance(col_1.loc[line], float):
                continue
            elif 'Project:' in col_1.loc[line]:
                starts.append(line)
            elif 'Min.' in col_1.loc[line]:
                ends.append(line)
        
        # Organizing the key cell info 
        targets = pd.DataFrame(data = {'starts': starts, 'ends': ends})
        
        # Creating the method scale final dataframe that will be filled with info
        # from the if loop
        output = pd.DataFrame()
        
        # Extracting the attribute 'sample'
        sample = mess.iloc[targets.iloc[0, 0]+6, 0]
        sample = sample.replace('ID: ', '')
        
        # The for loop iterates over the key cells to identify, extract, and
        # organize each sub-sheet bounded by the keys from targets. Each iteration
        # organizes one sub-sheet. At the end of the for loop, the organized
        # subsheet is appended to the output dataframe.
        for index in targets.index:
            start = targets.iloc[index, 0]
            end = targets.iloc[index, 1]
            site = mess.iloc[start+2, 0]
            site = site.replace('Site: Site of Interest ', '')
            columns = mess.iloc[start+14]
            temp_frame = mess.iloc[start+16: end-3]
            temp_frame = temp_frame.set_axis(labels = columns, axis = 1)
            if "In stats." in columns.values: # this raises an error if there are no XY coordinates
                temp_frame = temp_frame.drop(labels = 'In stats.', axis = 1)
            else:
                return "Error, coordinates missing"
            if '0' in temp_frame.columns:# if there are no trailing empty "0 columns" the code fails without this if loop
                temp_frame = temp_frame.drop(labels = '0', axis = 1)
            temp_frame['sample'] = [sample]*len(temp_frame.index)
            temp_frame['site'] = [site]*len(temp_frame.index)
            output = pd.concat(objs = [output, temp_frame], axis = 0, ignore_index = True)  
            
        # The rest of this is semi organizing the columns so that Sample, site, 
        # and spectrum are the first columns and Total is the last column. Also
        # removing NaN, because I CaN
        elms = output.columns
        elms = elms.drop(['Spectrum', 'Total', 'sample', 'site'])
        elms = list(elms)
        elms.append('Total')
        order = ['sample', 'site', 'Spectrum']  
        order.extend(elms)
        output = output[order]
        output = output.fillna('0')
        return output
        
    def change_coords(self, data, units_in = 'mm', units_out = 'm', x_column = 'X (mm)', y_column = 'Y (mm)', invert = True):
        '''
        A method that transforms coordinate data from the SMU SEM so that it can
        be properly used by a GIS tool, e.g. QGIS. Data must be organized in a
        sensible DataFrame, so probably call self.organize sheet first.

        Parameters
        ----------
        data : pandas DataFrame
            Organized DataFrame with spatial coordinates contained in columns.
        
        units_in : string = 'mm'
            The units of the input spatial data. Used to calculate the
            multiplication factoy. The default is 'mm'. Other options include
            "um" for micrometers. 
        
        units_out : string = 'm'
            The units that will be output by this method. The default is 'm'.
            In fact, the only currently accepted output unit is 'm'. This is only
            included to provide a platform for any possible future changes.
        
        x_column : string = 'X (mm)'
            The name of the column that contains the X coordinates. The default
            is 'X (mm)', which is the default name provided by the SMU SEM.
        
        y_column : string = 'Y (mm)'
            The name of the column that contains the Y coordinates. The default
            is 'Y (mm)', which is the default name provided by the SMU SEM
        
        invert : bool = True
            A boolean dictating whether or not the coordinates should be
            inverted. When moving SEM coordinates to a GIS, they must be inverted
            to make it work. As such, the default has been set to True.

        Returns
        -------
        output : pandas DataFrame
            The input dataframe with the origonal coordinates removed, and the
            transformed coordinates input. Please note, the new coordinate 
            columns are now the rightmost in the dataframe. 

        '''
        #import pandas as pd
        # Setting the transformation scale
        change = -0.001
        if units_in == 'mm' and units_out == 'm':
            change = -0.001
        elif units_in == 'um' and units_out == 'm':
            change = -0.000001
        else:
            print("Complain to Adam, cause this won't work")
        
        if invert == False:
            change *= -1
        
        output = data.copy()
        output["X ("+units_out+")"] = output[x_column].multiply(change)
        output["Y ("+units_out+")"] = output[y_column].multiply(change)
        output = output.drop(labels = [x_column, y_column], axis = 1)
        return output
    
    def batch_all(self):
        """
        Batch processes tif-hdr and .xlsx outputs. See self.to_tfw, self.organize_sheet
        and self.change_coords for more details
        """
        x = self.find_files(file_suffix='-tif.hdr')
        for name in x:
            self.to_tfw(hdr = name)
        y = self.find_files(file_suffix = '.xlsx')
        for name in y:
            spreadsheet = self.organize_sheet(frame = name)
            spreadsheet = self.change_coords(data = spreadsheet)
            spreadsheet.to_csv(self.path+"/transformed_coordinates.csv")
            break
        
    def batch_coords(self):
        """
        Batch processing of .xlsx outputs from the SEM. See self.organize_sheet, 
        and self.change_coordinates for details
        """
        y = self.find_files(file_suffix = '.xlsx')
        for name in y:
            spreadsheet = self.organize_sheet(frame = name)
            spreadsheet = self.change_coords(data = spreadsheet)
            spreadsheet.to_csv(self.path+"/transformed_coordinates.csv")
            break
    
def test_with_bar():
    test = mica(path = "Example_Data", output_path = "Output/With_bar")
    test.batch_tfw()
    test.batch_move()

def test_without_bar():
    test = mica(path = "Example_Data", output_path = "Output/Bar_removed")
    test.batch_tfw()
    test.batch_bar()

if __name__ == "__main__":
    test = mica(path = "Example_Data", output_path = "Output/Blank")
    test.batch_tfw()
    test.batch_bar()
