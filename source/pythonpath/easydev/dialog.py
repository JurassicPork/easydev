# coding: utf-8

import logging

import uno
import unohelper
from com.sun.star.awt import XItemListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import XFocusListener
from com.sun.star.awt import XActionListener

from com.sun.star.table import TableSortField
from com.sun.star.table.TableSortFieldType import AUTOMATIC

from org.universolibre.EasyDev import XLODialog
from org.universolibre.EasyDev import Macro
from easydev import comun
from easydev.tools import call_macro
from easydev.comun import LODefault
from easydev.setting import (
    LOG,
    NAME_EXT,
    COLORS,
    DECIMALS,
    FORMAT,
    DIALOGS,
)


log = logging.getLogger(NAME_EXT)


class ItemEvents(unohelper.Base, XItemListener):

    def __init__(self, dialog):
        self.dialog = dialog

    def disposing(self, event):
        pass

    def itemStateChanged(self, event):
        pass


class MapItemEvents(ItemEvents):

    def itemStateChanged(self, event):
        self.dialog.Model.Step = event.ItemId + 1
        return


class MouseEvents(unohelper.Base, XMouseListener):

    def __init__(self):
        pass

    def disposing(self, event):
        pass

    def mousePressed(self, event):
        pass

    def mouseReleased(self, event):
        pass

    def mouseEntered(self, event):
        pass

    def mouseExited(self, event):
        pass


class LinkMouseEvents(MouseEvents):

    def mouseEntered(self, event):
        obj = event.Source.Model
        obj.TextColor = COLORS['BLUE']
        return

    def mouseExited(self, event):
        obj = event.Source.Model
        obj.TextColor = 0
        return


class GridMouseEvents(MouseEvents):
    selected = False

    def mousePressed(self, event):
        obj = event.Source
        col = obj.getColumnAtPoint(event.X, event.Y)
        row = obj.getRowAtPoint(event.X, event.Y)
        if col == -1 and row == -1:
            if self.selected:
                obj.deselectAllRows()
            else:
                obj.selectAllRows()
            self.selected = not self.selected
        return


class GridDataFormMouseEvents(GridMouseEvents):

    def __init__(self, cell):
        self.cell = cell

    def mouseReleased(self, event):
        grid = event.Source
        col = grid.getColumnAtPoint(event.X, event.Y)
        row = grid.getRowAtPoint(event.X, event.Y)
        log.info('Col: {} Row: {}'.format(col, row))
        if col > -1 and row == -1:
            rango = comun.get_current_region(self.cell)
            field = TableSortField()
            field.Field = col
            field.IsAscending = True
            field.IsCaseSensitive = False
            field.FieldType = AUTOMATIC

            sort_des = rango.createSortDescriptor()
            sort_des[1].Name = "ContainsHeader"
            sort_des[1].Value = True
            sort_des[3].Name = "SortFields"
            sort_des[3].Value = uno.Any('[]com.sun.star.table.TableSortField', (field,))
            rango.sort(sort_des)
            log.info(rango.AbsoluteName)
        pass


class FocusEvents(unohelper.Base, XFocusListener):

    def disposing(self, event):
        pass

    def focusGained(self, event):
        pass

    def focusLost(self, event):
        pass


class ControlFocusEvents(unohelper.Base, XFocusListener):

    def __init__(self, color):
        self.color = color

    def focusGained(self, event):
        obj = event.Source.Model
        obj.Border = 0
        obj.BackgroundColor = self.color
        return

    def focusLost(self, event):
        obj = event.Source.Model
        obj.Border = 1
        obj.BackgroundColor = COLORS['WHITE']
        return


class ButtonEvents(unohelper.Base, XActionListener):

    def __init__(self, factory, macro):
        self.factory = factory
        self.macro = macro

    def disposing(self, event):
        pass

    def actionPerformed(self, event):
        control_name = '{}_action'.format(event.Source.Model.Name)
        if not self.macro.Name:
            self.macro.Name = control_name
        call_macro(self.factory, self.macro, (event,))
        return


class CloseButtonEvents(unohelper.Base, XActionListener):

    def __init__(self, dialog):
        self.dialog = dialog

    def disposing(self, event):
        pass

    def actionPerformed(self, event):
        self.dialog.endDialog(0)
        return


class LODialog(XLODialog, LODefault):
    decimals = DECIMALS
    numfmt = FORMAT.format(decimals)
    colorOnFocus = COLORS['YELLOW']

    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)
        self.factory = self._create_instance(
            'com.sun.star.script.provider.MasterScriptProviderFactory', False)

    def createDialog(self, data):
        """
            Create dialog from URL.
            path in OS or URI Specification vnd.sun.star.script
        """
        dp = self._create_instance('com.sun.star.awt.DialogProvider', True)
        if isinstance(data, Macro):
            if not data.Library:
                data.Library = 'Standard'
            data.Location = 'application'
            path = 'vnd.sun.star.script:{}.{}?location={}'.format(
                data.Library, data.Dialog, data.Location)
        elif comun.exists(data):
            path = comun.path_to_url(data)
        return dp.createDialog(path)
        #~ path_dialog = "vnd.sun.star.tdoc:/{}/Dialogs/{}/{}.xml".format(
            #~ path_current[1], module, name)

    def createControl(self, dialog, type_control, options):
        properties = comun.to_dict(options)
        base_properties = {
            'Width': 100,
            'Height': 12,
            'PositionX': 0,
            'PositionY': 0,
            'Step': 0,
            'TabIndex': 1,
        }
        controls = {
            'Button': 'com.sun.star.awt.UnoControlButtonModel',
            'CheckBox': 'com.sun.star.awt.UnoControlCheckBoxModel',
            'ComboBox': 'com.sun.star.awt.UnoControlComboBoxModel',
            'CurrencyField': 'com.sun.star.awt.UnoControlCurrencyFieldModel',
            'DateField': 'com.sun.star.awt.UnoControlDateFieldModel',
            'Edit': 'com.sun.star.awt.UnoControlEditModel',
            'FileControl': 'com.sun.star.awt.UnoControlFileControlModel',
            'FixedHyperlink': 'com.sun.star.awt.UnoControlFixedHyperlinkModel',
            'FixedLine': 'com.sun.star.awt.UnoControlFixedLineModel',
            'FixedText': 'com.sun.star.awt.UnoControlFixedTextModel',
            'FormattedField': 'com.sun.star.awt.UnoControlFormattedFieldModel',
            'GroupBox': 'com.sun.star.awt.UnoControlGroupBoxModel',
            'ImageControl': 'com.sun.star.awt.UnoControlImageControlModel',
            'ListBox': 'com.sun.star.awt.UnoControlListBoxModel',
            'NumericField': 'com.sun.star.awt.UnoControlNumericFieldModel',
            'PatternField': 'com.sun.star.awt.UnoControlPatternFieldModel',
            'ProgressBar': 'com.sun.star.awt.UnoControlProgressBarModel',
            'RadioButton': 'com.sun.star.awt.UnoControlRadioButtonModel',
            'ScrollBar': 'com.sun.star.awt.UnoControlScrollBarModel',
            'SimpleAnimation': 'com.sun.star.awt.UnoControlSimpleAnimationModel',
            'SpinButton': 'com.sun.star.awt.UnoControlSpinButtonModel',
            'Throbber': 'com.sun.star.awt.UnoControlThrobberModel',
            'TimeField': 'com.sun.star.awt.UnoControlTimeFieldModel',
            'Roadmap': 'com.sun.star.awt.UnoControlRoadmapModel',
            'Grid': 'com.sun.star.awt.grid.UnoControlGridModel',
        }
        controls_properties = {
            'FixedHyperlink': base_properties.copy(),
            'Roadmap': base_properties.copy(),
            'Grid': base_properties.copy(),
            'Edit': base_properties.copy(),
            'Button': base_properties.copy(),
            'DateField': base_properties.copy(),
        }
        controls_properties['Roadmap'].update({
            'Height': 100,
            'Text': 'Menu'}
        )
        controls_properties['Grid'].update({
            'Height': 100,
            'BackgroundColor': COLORS['WHITE'],
            'Sizeable': False,
            'ShowColumnHeader': True,
            'ShowRowHeader': True,
            'SelectionModel': 2,
            'UseGridLines': True}
        )
        controls_properties['Button'].update({
            'Label': 'CommandButton',
            'DefaultButton': False}
        )
        controls_properties['DateField'].update({
            'Dropdown': True}
        )

        #~ controls_properties['CheckBox'] = base_properties.update({
            #~ 'Label': 'CheckBox'})
        #~ controls_properties['ComboBox'] = base_properties.update({
            #~ 'Dropdown': True})
        #~ controls_properties['CurrencyField'] = base_properties.update({
            #~ 'Spin': True})
        #~ controls_properties['FileControl'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['FixedLine'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':5})
        #~ controls_properties['FixedText'] = base_properties.update({
            #~ 'Label':'Label'})
        #~ controls_properties['FormattedField'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['GroupBox'] = base_properties.update({
            #~ 'Width':100,
            #~ 'Height':30})
        #~ controls_properties['ImageControl'] = base_properties.update({
            #~ 'Width':30,
            #~ 'Height':30})
        #~ controls_properties['ListBox'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':30})
        #~ controls_properties['NumericField'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['PatternField'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['ProgressBar'] = base_properties.update({
            #~ 'Width':100,
            #~ 'Height':13})
        #~ controls_properties['RadioButton'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['ScrollBar'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['SimpleAnimation'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':30})
        #~ controls_properties['SpinButton'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})
        #~ controls_properties['Throbber'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':30})
        #~ controls_properties['TimeField'] = base_properties.update({
            #~ 'Width':60,
            #~ 'Height':13})

        """Create controls"""
        if not 'Name' in properties:
            return None
        dialog_model = dialog.getModel()
        # Only add if not exist
        if dialog_model.hasByName(properties['Name']):
            return None

        if type_control in controls:
            control = dialog_model.createInstance(controls[type_control])
            # Add default properties
            for k, v in controls_properties[type_control].items():
                if not k in properties:
                    properties[k] = controls_properties[type_control][k]
                if type_control == 'Grid':
                    properties['GridDataModel'] = self._create_instance(
                        'com.sun.star.awt.grid.DefaultGridDataModel')
                    properties['ColumnModel'] = self._add_columns_grid(properties['Columns'])
            # Only properties in control
            for k, v in properties.items():
                if control.getPropertySetInfo().hasPropertyByName(k):
                    # Properties special
                    #~ if propertie == 'StringItemList':
                        #~ uno.invoke(control, "setPropertyValue" , (
                            #~ "StringItemList", uno.Any(
                                #~ "[]string", properties[propertie])))
                    #~ else:
                    control.setPropertyValue(k, v)
        dialog_model.insertByName(properties['Name'], control)
        obj = self._listeners(dialog, type_control, control, properties)
        return obj

    def _listeners(self, dialog, type_control, control, properties):
        obj = dialog.getControl(properties['Name'])
        if type_control == 'Roadmap':
            if 'Options' in properties:
                self._add_options_roadmap(control, properties['Options'])
            obj.addItemListener(MapItemEvents(dialog))
        elif type_control == 'FixedHyperlink':
            obj.addMouseListener(LinkMouseEvents())
        elif type_control == 'Grid':
            if properties.get('DataForm', False):
                obj.addMouseListener(GridDataFormMouseEvents(properties['Cell']))
            else:
                obj.addMouseListener(GridMouseEvents())
        elif type_control == 'Edit':
            obj.addFocusListener(ControlFocusEvents(self.colorOnFocus))
        elif type_control == 'Button':
            macro = properties.get('Macro', False)
            obj.addActionListener(ButtonEvents(self.factory, macro))
        return obj

    def _add_options_roadmap(self, roadmap, options):
        for i, v in enumerate(options):
            opt = roadmap.createInstance()
            opt.ID = i
            opt.Label = v
            roadmap.insertByIndex(i, opt)
        return

    def _add_columns_grid(self, columns):
        #~ https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/XGridColumn.html
        column_model = self._create_instance(
            'com.sun.star.awt.grid.DefaultGridColumnModel')
        for col in columns:
            values = comun.to_dict(col)
            grid_column = self._create_instance('com.sun.star.awt.grid.GridColumn')
            for k, v in values.items():
                setattr(grid_column, k, v)
            column_model.addColumn(grid_column)
        return column_model

    def setGridData(self, grid, data, colFormat):
        grid_dm = grid.Model.GridDataModel
        grid_dm.removeAllRows()
        heading = tuple(range(1, len(data) + 1))
        if colFormat:
            rows = tuple(
                tuple(self._format(r, colFormat[i]) for i, r in enumerate(row)) for row in data
            )
        else:
            rows = tuple(tuple(self._format(r) for r in row) for row in data)
        grid_dm.addRows(heading, rows)

        rows = range(grid_dm.RowCount)
        colors = [COLORS['GRAY'] if r % 2 else COLORS['WHITE'] for r in rows]
        grid.Model.RowBackgroundColors = tuple(colors)
        return

    def _format(self, value, fmt=''):
        if fmt:
            return fmt.format(value)
        else:
            if isinstance(value, (int, float)):
                new_value = self.numfmt.format(value)
            else:
                new_value = value
        return new_value

    def setQuery(self, grid, query, colid):
        data = comun.parse_data_type(query)
        headers = data[0]
        row = data[1:2]
        rows = data[1:]
        col_fmt = False
        if colid:
            col_fmt = ('{}',) + ('',) * (len(headers) - 1)
        self._make_columns(grid, headers, row)
        self.setGridData(grid, rows, col_fmt)
        return

    def _make_columns(self, grid, headers, row):
        if row:
            align = tuple(self._get_align(r) for r in row[0])
        else:
            align = (0,) * len(headers)
        width = (grid.Model.Width - 30) / len(headers)
        columns = []
        for i, v in enumerate(headers):
            col = {}
            col['Title'] = v
            col['HorizontalAlign'] = align[i]
            col['ColumnWidth'] = width
            columns.append(col)
        self.setGridColumns(grid, columns)
        return

    def _get_align(self, value):
        align = 0
        if isinstance(value, (int, float)):
            align = 2
        return align

    def setGridColumns(self, grid, columns):
        columns_model = grid.Model.ColumnModel
        columns_model.setDefaultColumns(len(columns))
        for i, col in enumerate(columns):
            column = columns_model.getColumn(i)
            for k, v in col.items():
                setattr(column, k, v)
        return

    def getGridData(self, grid, exclude):
        gdm = grid.Model.GridDataModel
        gcm = grid.Model.ColumnModel
        cell = getattr(gdm, 'getCellData')
        cols = range(gdm.ColumnCount)
        rows = range(gdm.RowCount)
        data = tuple(
            tuple(self._cell(cell, gcm, c, r) for c in cols if not c in exclude) for r in rows
        )
        return data

    def _cell(self, cell, gcm, c, r):
        value = cell(c, r)
        if gcm.getColumn(c).Identifier:
            try:
                value = float(value.strip('$').replace(',', ''))
            except ValueError:
                value = cell(c, r)
        return value

    def gridDataForm(self, cell):
        try:
            dlg = self._get_dlg()
            data = comun.get_data_range(cell)
            grid = self._create_grid_data_form(dlg, cell)
            self._set_grid_data_form(grid, data)
            self._data_form_listeners(dlg)
            dlg.execute()
            dlg.dispose()
        except:
            log.error('Data form', exc_info=True)
        return

    def _data_form_listeners(self, dlg):
        cmd_close = dlg.getControl('cmd_close')
        cmd_close.addActionListener(CloseButtonEvents(dlg))
        txt_search = dlg.getControl('txt_search')
        txt_search.addFocusListener(ControlFocusEvents(self.colorOnFocus))
        return

    def _get_dlg(self):
        path_dlg = comun.get_path_dlg(DIALOGS['DATAFORM'])
        dlg = self.createDialog(path_dlg)
        comun.set_icons(dlg.Model)
        return dlg

    def _create_grid_data_form(self, dialog, cell):
        dm = dialog.Model
        properties = {
            'Name': 'grid',
            'Width': dm.Width - 10,
            'Height': dm.Height - 44,
            'PositionX': 5,
            'PositionY': 22,
            'Step': 0,
            'TabIndex': 1,
            'Columns': (),
            'DataForm': True,
            'Cell': cell,
        }
        grid = self.createControl(dialog, 'Grid', properties)
        return grid

    def _set_grid_data_form(self, grid, data):
        headers = data[0]
        row = data[1:2]
        rows = data[1:]
        self._make_columns(grid, headers, row)
        col_fmt = ('{}',) * len(headers)
        self.setGridData(grid, rows, col_fmt)
        return
        #~ for control in self.dialog.Controls:
            #~ if control.ImplementationName in EXCLUDE_CONTROLS:
                #~ continue
            #~ setattr(self, control.Model.Name, control)
