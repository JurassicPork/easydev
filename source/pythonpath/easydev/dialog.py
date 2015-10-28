# coding: utf-8

import logging

import unohelper
from com.sun.star.awt import XItemListener
from com.sun.star.awt import XMouseListener

from org.universolibre.EasyDev import XLODialog
from easydev import comun
from easydev.comun import LODefault
from easydev.setting import LOG, NAME_EXT, COLORS, DECIMALS, FORMAT


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


class LODialog(XLODialog, LODefault):
    decimals = DECIMALS
    numfmt = FORMAT.format(decimals)

    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)

    def createDialog(self, path):
        """Create dialog from URL."""
        dp = self._create_instance('com.sun.star.awt.DialogProvider', True)
        return dp.createDialog(comun.path_to_url(path))

    def createControl(self, dialog, type_control, options):
        properties = comun.to_dict(options)
        base_properties = {
            'Width': 100,
            'Height': 10,
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
        }
        controls_properties['Roadmap'].update({
            'Height': 100,
            'Text': 'Menu'})
        controls_properties['Grid'].update({
            'Height': 100,
            'BackgroundColor': COLORS['WHITE'],
            'Sizeable': False,
            'ShowColumnHeader': True,
            'ShowRowHeader': True,
            'SelectionModel': 2,
            'UseGridLines': True})

        controls_properties['Button'] = base_properties.update({
            'Label': 'CommandButton',
            'DefaultButton': False,
            'PushButtonType': 0})
        controls_properties['CheckBox'] = base_properties.update({
            'Label': 'CheckBox'})
        controls_properties['ComboBox'] = base_properties.update({
            'Dropdown': True})
        controls_properties['CurrencyField'] = base_properties.update({
            'Spin': True})
        controls_properties['DateField'] = base_properties.update({
            'Dropdown':True})
        controls_properties['Edit'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['FileControl'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['FixedLine'] = base_properties.update({
            'Width':60,
            'Height':5})
        controls_properties['FixedText'] = base_properties.update({
            'Label':'Label'})
        controls_properties['FormattedField'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['GroupBox'] = base_properties.update({
            'Width':100,
            'Height':30})
        controls_properties['ImageControl'] = base_properties.update({
            'Width':30,
            'Height':30})
        controls_properties['ListBox'] = base_properties.update({
            'Width':60,
            'Height':30})
        controls_properties['NumericField'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['PatternField'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['ProgressBar'] = base_properties.update({
            'Width':100,
            'Height':13})
        controls_properties['RadioButton'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['ScrollBar'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['SimpleAnimation'] = base_properties.update({
            'Width':60,
            'Height':30})
        controls_properties['SpinButton'] = base_properties.update({
            'Width':60,
            'Height':13})
        controls_properties['Throbber'] = base_properties.update({
            'Width':60,
            'Height':30})
        controls_properties['TimeField'] = base_properties.update({
            'Width':60,
            'Height':13})

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
            obj.addMouseListener(GridMouseEvents())
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
            rows = tuple(tuple(self._format(r, colFormat[i]) for i, r in enumerate(row)) for row in data)
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



