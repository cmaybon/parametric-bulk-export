from . import commands
from .lib import fusion360utils as futil
import adsk.core
import adsk.fusion
import traceback
import os

COMMAND_NAME = "Parametric Export"
COMMAND_DESCRIPTION = "Bulk export meshes, changing selected parameters." 
COMMAND_RESOURCES = ""
COMMAND_ID = "parametric-bulk-export"
TARGET_WORKSPACE = "FusionSolidEnvironment"
TARGET_PANEL = "SolidScriptsAddinsPanel"
_handlers = []

CACHED_VARIATION_DATA = dict()


class ParametricBulkExporter:
    def __init__(self) -> None:
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
        self.design: adsk.fusion.Design = self.app.activeProduct
        self.original_user_parameter_data = dict()
    
    @property
    def active_component(self):
        return self.design.activeComponent

    def output_folder_dialog(self):
        folder_dialog = self.ui.createFolderDialog()
        folder_dialog.title = "Select output folder"
        result = folder_dialog.showDialog()
        if result != adsk.core.DialogResults.DialogOK:
            return None
        return folder_dialog.folder
    
    def cache_user_parameter_values(self):
        for parameter in self.design.userParameters:
            self.original_user_parameter_data[parameter.name] = parameter.expression

    def restore_user_parameters_from_cache(self):
        for name, value in self.original_user_parameter_data.items():
            self.design.userParameters.itemByName(name).expression = value

    def get_table_string_input(self, table, row, column):
        return table.getInputAtPosition(row, column)

    def apply_user_parameter_changes(self, change_data):
        for name, new_value in change_data.items():
            self.design.userParameters.itemByName(name).expression = new_value

    def create_file_name(self, base_filename, changed_parameters):
        for name, value in changed_parameters.items():
            base_filename += f"_{str(value).replace(' ', '')}_{name}"
        return base_filename

    def export(self, parameter_table, do_stl, do_step, do_obj):
        output_folder = self.output_folder_dialog()
        self.cache_user_parameter_values()
        futil.log("user parameters cached")
        for column in range(parameter_table.numberOfColumns):
            # Skip the user parameter name and value columns
            if column in [0, 1]:
                continue
            futil.log(f"exporting variation {column - 1}")
            changed_parameters = dict()
            for row in range(parameter_table.rowCount):
                # Skip the header row
                if row == 0:
                    continue

                parameter_for_row = self.get_table_string_input(parameter_table, row, 0)
                cell_input = self.get_table_string_input(parameter_table, row, column)
                if cell_input is None or bool(cell_input.value) is False:
                    continue
                changed_parameters[str(parameter_for_row.value)] = cell_input.value
            if len(changed_parameters) == 0:
                futil.log(f"not exporting variation {column - 1}, no parameters were changed")
                continue
            self.apply_user_parameter_changes(changed_parameters)
            self.export_meshes(output_folder,
                               self.create_file_name("testing", changed_parameters),
                               self.active_component,
                               do_stl,
                               do_step,
                               do_obj)
        self.restore_user_parameters_from_cache()
        futil.log("user parameters restored from cache")

    def export_meshes(self, output_folder, file_name, component, do_stl, do_step, do_obj):
        export_manager = component.parentDesign.exportManager
        output_path = os.path.join(output_folder, file_name)
        if do_stl:
            futil.log("exporting stl")
            options = export_manager.createSTLExportOptions(component, output_path + ".stl")
            export_manager.execute(options)
        if do_step:
            futil.log("exporting step")
            options = export_manager.createSTEPExportOptions(output_path, component)
            export_manager.execute(options)
        if do_obj:
            futil.log("exporting obj")
            options = export_manager.createOBJExportOptions(component, output_path)
            export_manager.execute(options)


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface

    def notify(self, args):
        try:
            self._notify(args)
        except Exception:
            if self.ui:
                self.ui.messageBox(f"command executed failed:\n{traceback.format_exc()}")

    def _notify(self, args):
        inputs = args.command.commandInputs

        parameter_table = inputs.itemById("parameterBulkTable")
        do_stl = inputs.itemById("exportStlMeshBool").value
        do_step = inputs.itemById("exportStepMeshBool").value
        do_obj = inputs.itemById("exportObjMeshBool").value
        ParametricBulkExporter().export(parameter_table, do_stl, do_step, do_obj)


class CommandDeactivateHandler(adsk.core.CommandEventHandler):
    def __init__(self, table_input):
        super().__init__()
        self.table_input = table_input

    def notify(self, args) -> None:
        CACHED_VARIATION_DATA.clear()
        futil.log("caching variation data")
        for column in range(self.table_input.numberOfColumns):
            if column in [0, 1]:
                continue
            for row in range(self.table_input.rowCount):
                if row == 0:
                    continue
                parameter_cell = self.table_input.getInputAtPosition(row, 0)
                variation_cell = self.table_input.getInputAtPosition(row, column)
                cache_key = f"variation{column - 1}"
                if CACHED_VARIATION_DATA.get(cache_key) is None:
                    CACHED_VARIATION_DATA[cache_key] = dict()
                CACHED_VARIATION_DATA[cache_key][parameter_cell.value] = variation_cell.value
        futil.log("finished caching variation data")


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
        self.design: adsk.fusion.Design = self.app.activeProduct
        self.export_variations_count = 10
        self.total_columns = 2 + self.export_variations_count

    def notify(self, args):
        try:
            self._notify(args)
        except Exception:
            if self.ui:
                self.ui.messageBox(f"Panel command created failed:\n{traceback.format_exc()}")

    def _notify(self, args):
        cmd = args.command
        on_execute = CommandExecuteHandler()
        cmd.execute.add(on_execute)
        _handlers.append(on_execute)
        inputs = cmd.commandInputs

        export_file_types_group = inputs.addGroupCommandInput("exportFileTypes", "File Types")
        export_file_types_group.children.addBoolValueInput("exportStlMeshBool", "STL", True)
        export_file_types_group.children.addBoolValueInput("exportStepMeshBool", "Step", True)
        export_file_types_group.children.addBoolValueInput("exportObjMeshBool", "Obj", True)
        table = self.create_parameter_table(inputs)

        on_deactivate = CommandDeactivateHandler(table)
        cmd.deactivate.add(on_deactivate)
        _handlers.append(on_deactivate)

    def create_parameter_table(self, cmd_inputs):
        table_input = cmd_inputs.addTableCommandInput("parameterBulkTable",
                                                      "Param Changes",
                                                      self.total_columns,
                                                      "3:2:1")
        table_input.columnSpacing = 0
        table_input.maximumVisibleRows = self.total_columns
        table_input.tablePresentationStyle = 2

        self.add_header_row(table_input, cmd_inputs)
        self.add_parameter_rows(table_input, cmd_inputs)
        return table_input

    def add_header_row(self, table_input, cmd_inputs):
        def _get_style_string(text):
            return f'<div align="center", style="font-size:12px"><b>{text}</b></div>'

        def _add_textbox_command_input(column_index, text):
            name = f"col{column_index}Header"
            header = cmd_inputs.addTextBoxCommandInput(name, name, _get_style_string(text), 1, True)
            table_input.addCommandInput(header, 0, column_index)

        _add_textbox_command_input(0, "Parameter Name")
        _add_textbox_command_input(1, "Value")
        for i in range(self.export_variations_count):
            _add_textbox_command_input(2 + i, f"Export {i + 1}")

    def add_parameter_rows(self, table_input, cmd_inputs):
        def _add_string_command_input(row, column, name, value):
            cell = cmd_inputs.addStringValueInput(f"{name}{column}", name, value)
            table_input.addCommandInput(cell, row, column)

        current_row = 1
        for parameter in self.design.userParameters:
            _add_string_command_input(current_row, 0, parameter.name, parameter.name)
            _add_string_command_input(current_row, 1, parameter.name, parameter.expression)

            for i in range(self.export_variations_count):
                cached_value = CACHED_VARIATION_DATA.get(f"variation{i + 1}")
                if cached_value is not None:
                    cached_value = cached_value.get(parameter.name)
                if cached_value is None:
                    cached_value = ""
                parameter_variation_input = cmd_inputs.addStringValueInput(f"variation{i}Input", f"Variation {i}", cached_value)
                table_input.addCommandInput(parameter_variation_input, current_row, 2 + i)
            current_row += 1


def get_add_in_command_definition(ui):
    command_definitions = ui.commandDefinitions
    command_definition = command_definitions.itemById(COMMAND_ID)
    if not command_definition:
        command_definition = command_definitions.addButtonDefinition(COMMAND_ID,
                                                                     COMMAND_NAME,
                                                                     COMMAND_DESCRIPTION,
                                                                     COMMAND_RESOURCES)
    return command_definition


def command_control_by_id_for_panel(command_id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not command_id:
        ui.messageBox('commandControl id is not specified')
        return None
    workspaces = ui.workspaces
    modeling_workspace = workspaces.itemById(TARGET_WORKSPACE)
    toolbar_panels = modeling_workspace.toolbarPanels
    toolbar_panel = toolbar_panels.itemById(TARGET_PANEL)
    toolbar_controls = toolbar_panel.controls
    toolbar_control = toolbar_controls.itemById(command_id)
    return toolbar_control


def command_definition_by_id(command_id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not command_id:
        ui.messageBox('command_definition id is not specified')
        return None
    command_definitions = ui.commandDefinitions
    command_definition = command_definitions.itemById(command_id)
    return command_definition


def destroy_object(ui_obj, to_be_delete_obj):
    if not ui_obj and not to_be_delete_obj:
        return

    if to_be_delete_obj.isValid:
        to_be_delete_obj.deleteMe()
    else:
        ui_obj.messageBox('toBeDeleteObj is not a valid object')


def run(context):
    ui = None

    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        start_add_in(context, app, ui)
    except Exception:
        if ui:
            ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))
        futil.handle_error('run')


def start_add_in(context, app, ui):
    command_definition = get_add_in_command_definition(ui)
    on_command_created = CommandCreatedHandler()
    command_definition.commandCreated.add(on_command_created)
    _handlers.append(on_command_created)

    workspaces = ui.workspaces
    modeling_workspace = workspaces.itemById(TARGET_WORKSPACE)
    toolbar_panels = modeling_workspace.toolbarPanels
    toolbar_panel = toolbar_panels.itemById(TARGET_PANEL)
    toolbar_controls_panel = toolbar_panel.controls
    toolbar_control_panel = toolbar_controls_panel.itemById(COMMAND_ID)
    if not toolbar_control_panel:
        toolbar_control_panel = toolbar_controls_panel.addCommand(command_definition, "")
        toolbar_control_panel.isVisible = True
        futil.log(f"{COMMAND_ID} successfully added to add ins panel")
    commands.start()


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        obj_array = []

        command_control_panel = command_control_by_id_for_panel(COMMAND_ID)
        if command_control_panel:
            obj_array.append(command_control_panel)

        command_definition = command_definition_by_id(COMMAND_ID)
        if command_definition:
            obj_array.append(command_definition)

        for obj in obj_array:
            destroy_object(ui, obj)

        futil.clear_handlers()
        commands.stop()
    except Exception:
        futil.handle_error('stop')
