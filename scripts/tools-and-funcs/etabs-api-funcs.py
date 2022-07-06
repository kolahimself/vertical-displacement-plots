import os 
import sys
import comtypes.client 


def connect_to_etabs():
    """
    
    Attaching to a Manually Started Instance of ETABS 
    
    Returns:
    SapModel: type cOAPI pointer
    """

    # Create API helper object
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    
    # Attach to a running instance of ETABS
    try:
        # Get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject") 
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
    
    # Create SapModel object
    SapModel = myETABSObject.SapModel
    
    return SapModel


def get_story_elevations(SapModel):
    """
    Retrieves the story elevations of a tower in the model.
    
    Parameters:
        :param SapModel: cOAPI indicator
    
    Returns:
    a list of the retrieved story elevations
    
    story_elevations: list | float | int 
    """
    
    # Get the data using API
    story_info = SapModel.Story.GetStories_2()
    
    # Extract the story elevations into a list
    story_elevations = list(story_info[3])
    
    return story_elevations;


def get_combinations(SapModel):
    """
    Retrieves the available load Combinations into a list.
    
    Parameters:
        :param SapModel: cOAPI indicator
        
    Returns:
    A list of load Combinations
    
    combos : list
    """
    
    combos = list(SapModel.RespCombo.GetNameList()[1])
    
    return combos;


def jointDisp_export(SapModel, savepath):
    """
    Exports a dataset of joint displacements from ETABS into a spreadsheet file in a desired path.
    
    Parameters:
        :param SapModel: cOAPI indicator
        :param savepath | string, the specified path as to which the spreadsheet file is to be saved
    
    Returns:
    Returns spreadsheet file to the specified directory
    
    Exception:
    Function raises an exception when 'Joint Displacement' table is absent in ETABS, ensure that analysis has been run.
    
    Examples:
    >> jointDisp_export(SapModel, savepath = 'disp-values.csv')
    Verify file existence at desktop, ✅
    
    >>  jointDisp_export(SapModel, savepath = 'disp-values.xlsx')
    Verify file existence at desktop, ✅
    """
    
    # Retrieve the table key with API (i.e 'Joint Displacement')
    key = SapModel.DatabaseTables.GetAvailableTables()[2][62]
    
    # Retrieve all the fields in the table (e.g ['Story', 'Label', 'Unique Name'..])
    fieldList = list(SapModel.DatabaseTables.GetAllFieldsInTable(TableKey = ti)[2])
    
    # OPTIONAL unless you want these fields removed.
    # unwanted_fields = ['StepType', 'StepNumber', 'StepLabel', 'Rx', 'Ry', 'Rz', 'Ux', 'Uy']
    # fieldList = [field for field in fieldList if field not in unwanted_fields]
    
    # Export the spreadsheet file to the desired path
    SapModel.DatabaseTables.GetTableForDisplayCSVFile(TableKey = key,
                                                      FieldKeyList = fieldList,
                                                      GroupName = 'All',
                                                      csvFilePath = savepath);
