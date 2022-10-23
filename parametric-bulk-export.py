from sre_constants import OP_IGNORE
from . import commands
from .lib import fusion360utils as futil
import adsk.core
import adsk.fusion
import traceback
import os

COMMAND_NAME = "Parametrid Export"
COMMAND_DESCRIPTION = "Bulk export meshes, changing selected parameters." 
COMMAND_RESOURCES = ""
COMMAND_ID = "parametric-bulk-export"
TARGET_WORKSPACE = "FusionSolidEnvironment"
TARGET_PANNEL = "SolidScriptsAddinsPanel"

# global set of event handlers to keep them referenced for the duration of the command
_handlers = []


class ParametricBulkExporter:
    def __init__(self) -> None:
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
        self.design = self.app.activeProduct
        self.original_user_paramerter_data = dict()
    
    @property
    def activeComponent(self):
        return self.design.activeComponent

    def outputFolderDialog(self):
        folder_dialog = self.ui.createFolderDialog()
        folder_dialog.title = "Select output folder"
        result = folder_dialog.showDialog()
        if result != adsk.core.DialogResults.DialogOK:
            return None
        return folder_dialog.folder
    
    def cacheUserParameterValues(self):
        for parameter in self.design.userParameters:
            self.original_user_paramerter_data[parameter.name] = parameter.expression

    def restoreUserParametersFromCache(self):
        for name, value in self.original_user_paramerter_data.items():
            self.design.userParameters.itemByName(name).expression = value

    def getTableStringInput(self, table, row, column):
        return table.getInputAtPosition(row, column)

    def applyUserParameterChanges(self, change_data):
        for name, new_value in change_data.items():
            self.design.userParameters.itemByName(name).expression = new_value

    def createFileName(self, base_filename, changed_parameters):
        for name, value in changed_parameters.items():
            base_filename += f"_{str(value).replace(' ', '')}_{name}"
        return base_filename

    def export(self, parameter_table, do_stl, do_step, do_obj):
        output_folder = self.outputFolderDialog()
        self.cacheUserParameterValues()
        for column in range(parameter_table.numberOfColumns):
            if column in [0, 1]:
                continue
            changed_parameters = dict()
            for row in range(parameter_table.rowCount):
                if row == 0:
                    continue
                cell_input = self.getTableStringInput(parameter_table, row, column)
                if cell_input is None or bool(cell_input.value) is False:
                    continue
                changed_parameters[str(self.getTableStringInput(parameter_table, row, 0).value)] = cell_input.value
            if len(changed_parameters) == 0:
                continue
            self.applyUserParameterChanges(changed_parameters)
            self.exportMeshes(output_folder, self.createFileName("testing", changed_parameters), self.activeComponent, do_stl, do_step, do_obj)
        self.restoreUserParametersFromCache()

    def exportMeshes(self, output_folder, file_name, component, do_stl, do_step, do_obj):
        export_manager = component.parentDesign.exportManager
        output_path = os.path.join(output_folder, file_name)
        if do_stl:
            options = export_manager.createSTLExportOptions(component, output_path + ".stl")
            export_manager.execute(options)
        if do_step:
            options = export_manager.createSTEPExportOptions(output_path, component)
            export_manager.execute(options)
        if do_obj:
            options = export_manager.createOBJExportOptions(component, output_path)
            export_manager.execute(options)


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
    max_export_iterations = 10
    all_parameters = getAllUserParameters()
    tableInput = cmdInputs.addTableCommandInput("parameterBulkTable", "Param Changes", 2 + max_export_iterations, "3:2:1")
    tableInput.columnSpacing = 0
    tableInput.maximumVisibleRows = 12
    tableInput.tablePresentationStyle = 2

    parameterNameColumnHeader = cmdInputs.addTextBoxCommandInput("col0Header", "col 0 Header", '<div align="center", style="font-size:12px"><b>Parameter Name</b></div>', 1, True)
    parameterExpressionColumnHeader = cmdInputs.addTextBoxCommandInput("col1Header", "col 1 Header", '<div align="center", style="font-size:12px"><b>Value</b></div>', 1, True)
    tableInput.addCommandInput(parameterNameColumnHeader, 0, 0)
    tableInput.addCommandInput(parameterExpressionColumnHeader, 0, 1)
    for i in range(max_export_iterations):
        exportVariationColumnHeader = cmdInputs.addTextBoxCommandInput(f"col{2 + i}Header", f"col {2 + i} Header", f'<div align="center", style="font-size:12px"><b>Export {i + 1}</b></div>', 1, True)
        tableInput.addCommandInput(exportVariationColumnHeader, 0, 2 + i)
    
    current_row = 1
    for parameter in all_parameters:
        parameter_name_textbox = cmdInputs.addStringValueInput(f"{parameter.name}TextBox", parameter.name, parameter.name)
        parameter_value_textbox = cmdInputs.addStringValueInput(f"{parameter.name}Value", parameter.name, parameter.expression) 
        tableInput.addCommandInput(parameter_name_textbox, current_row, 0)
        tableInput.addCommandInput(parameter_value_textbox, current_row, 1)
        
        for i in range(max_export_iterations):
            parameter_variation_input = cmdInputs.addStringValueInput(f"variation{i}Input", "Varation {i}", "")
            tableInput.addCommandInput(parameter_variation_input, current_row, 2 + i)
        current_row += 1
    return tableInput


def run(context):
    ui = None

    class CommandExecuteHandler(adsk.core.CommandEventHandler):
        def __init__(self):
            super().__init__()
        def notify(self, args):
            try:
                print("running command...")
                inputs = args.command.commandInputs

                parameter_table = inputs.itemById("parameterBulkTable")
                do_stl = inputs.itemById("exportStlMeshBool").value
                do_step = inputs.itemById("exportStepMeshBool").value
                do_obj = inputs.itemById("exportObjMeshBool").value
                ParametricBulkExporter().export(parameter_table, do_stl, do_step, do_obj)
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