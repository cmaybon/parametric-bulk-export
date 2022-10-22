from . import commands
from .lib import fusion360utils as futil
import adsk.core
import adsk.fusion
import traceback

COMMAND_NAME = "Parametrid Export"
COMMAND_DESCRIPTION = "Bulk export meshes, changing selected parameters." 
COMMAND_RESOURCES = "./resources"
COMMAND_ID = "parametric-bulk-export"
TARGET_WORKSPACE = "FusionSolidEnvironment"
TARGET_PANNEL = "SolidScriptsAddinsPanel"

# global set of event handlers to keep them referenced for the duration of the command
_handlers = []


def commandControlByIdForPanel(command_id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not command_id:
        ui.messageBox('commandControl id is not specified')
        return None
    workspaces = ui.workspaces
    modelingWorkspace = workspaces.itemById(TARGET_WORKSPACE)
    toolbarPanels = modelingWorkspace.toolbarPanels
    toolbarPanel = toolbarPanels.itemById(TARGET_PANNEL)
    toolbarControls = toolbarPanel.controls
    toolbarControl = toolbarControls.itemById(command_id)
    return toolbarControl


def commandDefinitionById(command_id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not command_id:
        ui.messageBox('commandDefinition id is not specified')
        return None
    commandDefinitions = ui.commandDefinitions
    commandDefinition = commandDefinitions.itemById(command_id)
    return commandDefinition


def destroyObject(uiObj, toBeDeleteObj):
    if uiObj and toBeDeleteObj:
        if toBeDeleteObj.isValid:
            toBeDeleteObj.deleteMe()
        else:
            uiObj.messageBox('toBeDeleteObj is not a valid object')


def getAllUserParameters():
    app = adsk.core.Application.get()
    design = app.activeProduct
    return design.userParameters


def createParameterTable(cmdInputs):
    all_parameters = getAllUserParameters()
    tableInput = cmdInputs.addTableCommandInput("parameterBulkTable", "Param Changes", 3, "3:2:1")
    tableInput.columnSpacing = 0
    tableInput.maximumVisibleRows = 12
    tableInput.tablePresentationStyle = 2

    parameterNameColumnHeader = cmdInputs.addTextBoxCommandInput("col0Header", "col 0 Header", '<div align="center", style="font-size:12px"><b>Parameter Name</b></div>', 1, True)
    parameterExpressionColumnHeader = cmdInputs.addTextBoxCommandInput("col1Header", "col 1 Header", '<div align="center", style="font-size:12px"><b>Value</b></div>', 1, True)
    export1VariationColumnHeader = cmdInputs.addTextBoxCommandInput("col2Header", "col 2 Header", '<div align="center", style="font-size:12px"><b>Export 1</b></div>', 1, True)
    tableInput.addCommandInput(parameterNameColumnHeader, 0, 0)
    tableInput.addCommandInput(parameterExpressionColumnHeader, 0, 1)
    tableInput.addCommandInput(export1VariationColumnHeader, 0, 2)

    current_row = 1
    for parameter in all_parameters:
        parameter_name_textbox = cmdInputs.addStringValueInput(f"{parameter.name}TextBox", parameter.name, parameter.name)
        parameter_value_textbox = cmdInputs.addStringValueInput(f"{parameter.name}Value", parameter.name, parameter.expression)
        tableInput.addCommandInput(parameter_name_textbox, current_row, 0)
        tableInput.addCommandInput(parameter_value_textbox, current_row, 1)
        current_row += 1


def run(context):
    ui = None

    class CommandExecuteHandler(adsk.core.CommandEventHandler):
        def __init__(self):
            super().__init__()
        def notify(self, args):
            try:
                print("running command...")
            except:
                if ui:
                    ui.messageBox('command executed failed:\n{}'.format(traceback.format_exc()))

    class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
        def __init__(self):
            super().__init__() 
        def notify(self, args):
            try:
                cmd = args.command
                onExecute = CommandExecuteHandler()
                cmd.execute.add(onExecute)
                # keep the handler referenced beyond this function
                _handlers.append(onExecute)

                inputs = cmd.commandInputs

                export_file_types_group = inputs.addGroupCommandInput("exportFileTypes", "File Types")
                export_stl_checkbox_input = export_file_types_group.children.addBoolValueInput("exportStlMeshBool", "STL", True)
                export_step_checkbox_input = export_file_types_group.children.addBoolValueInput("exportStepMeshBool", "Step", True)
                export_obj_checkbox_input = export_file_types_group.children.addBoolValueInput("exportObjMeshBool", "Obj", True)

                createParameterTable(inputs)
            except:
                if ui:
                    ui.messageBox('Panel command created failed:\n{}'.format(traceback.format_exc()))

    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        commandDefinitions = ui.commandDefinitions
        commandDefinition = commandDefinitions.itemById(COMMAND_ID)
        if not commandDefinition:
            commandDefinition = commandDefinitions.addButtonDefinition(COMMAND_ID, COMMAND_NAME, COMMAND_DESCRIPTION, COMMAND_RESOURCES)

        onCommandCreated = CommandCreatedHandler()
        commandDefinition.commandCreated.add(onCommandCreated)
        _handlers.append(onCommandCreated)

        workspaces = ui.workspaces
        modelingWorkspace = workspaces.itemById(TARGET_WORKSPACE)
        toolbarPanels = modelingWorkspace.toolbarPanels
        toolbarPanel = toolbarPanels.itemById(TARGET_PANNEL)
        toolbarControlsPannel = toolbarPanel.controls
        toolbarControlPannel = toolbarControlsPannel.itemById(COMMAND_ID)
        if not toolbarControlPannel:
            toolbarControlPannel = toolbarControlsPannel.addCommand(commandDefinition, "")
            toolbarControlPannel.isVisible = True
            print(f"{COMMAND_ID} successfully added to add ins pannel")


        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.start()

    except:
        if ui:
            ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))
        futil.handle_error('run')


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        objArray = []

        commandControlPanel = commandControlByIdForPanel(COMMAND_ID)
        if commandControlPanel:
            objArray.append(commandControlPanel)

        commandDefinition = commandDefinitionById(COMMAND_ID)
        if commandDefinition:
            objArray.append(commandDefinition)

        for obj in objArray:
            destroyObject(ui, obj)

        # Remove all of the event handlers your app has created
        futil.clear_handlers()

        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.stop()

    except:
        futil.handle_error('stop')