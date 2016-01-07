# coding: utf-8

import logging
from easydev.loapp import LOApp
from org.universolibre.EasyDev import XLOChart
from easydev.setting import LOG, NAME_EXT, PY2
from easydev import comun


if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)


class LOChart(XLOChart, LOApp):

    def __init__(self, ctx, sm, desktop, toolkit):
        LOApp.__init__(self, ctx, sm, desktop, toolkit)

    def chartAdd(self, data):
        charts = data.Sheet.getCharts()
        charts.addNewByName(data.Name, data.PosSize, (), True, True)
        chart = charts.getByName(data.Name).getEmbeddedObject()
        chart.setDiagram(chart.createInstance('com.sun.star.chart.{}'.format(data.Type)))
        self._add_series(chart, data.Doc, data.Series)

        properties = comun.to_dict(data.Properties)
        for k, v in properties.items():
            if k == 'Title':
                chart.getTitle().String = v
                continue
            if hasattr(chart, k):
                setattr(chart, k, v)

        return chart

    def _add_series(self, chart, doc, series):
        if not series:
            return
        dp = doc.createInstance('com.sun.star.chart2.data.DataProvider')
        chart.attachDataProvider(dp)
        chart_type = chart.getFirstDiagram().getCoordinateSystems()[0].getChartTypes()[0]
        new_series = []
        for serie in series:
            new_serie = self._create_serie(dp, serie)
            new_series.append(new_serie)
        chart_type.setDataSeries(tuple(new_series))
        return

    def _create_serie(self, dp, data):
        serie = self._create_instance('com.sun.star.chart2.DataSeries', False)
        properties = comun.to_dict(data.Properties)
        range_y = data.Y
        if not isinstance(range_y, str):
            range_y = range_y.AbsoluteName
        range_x = data.X
        if not isinstance(range_x, str):
            range_x = range_x.AbsoluteName
        range_title = data.Title
        if not isinstance(range_title, str):
            range_title = range_title.AbsoluteName

        data_y = self._create_instance('com.sun.star.chart2.data.LabeledDataSequence')
        values_y = self._create_data(dp, range_y, 'values-y')
        data_y.setValues(values_y)
        if data.Title:
            label = self._create_data(dp, range_title, '')
            data_y.setLabel(label)

        data_x = self._create_instance('com.sun.star.chart2.data.LabeledDataSequence')
        values_x = self._create_data(dp, range_x, 'values-x')
        data_x.setValues(values_x)

        serie.setData((data_y, data_x))
        for k, v in properties.items():
            if hasattr(serie, k):
                setattr(serie, k, v)
        return serie

    def _create_data(self, dp, rango, role):
        data = dp.createDataSequenceByRangeRepresentation(rango)
        if not data is None:
            data.Role = role
        return data


