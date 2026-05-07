# -*- coding: utf-8 -*-
from __future__ import division, print_function

__title__ = 'Ductulator'
__author__ = 'GMoreno'
__doc__ = 'ENVAR duct sizing reference with round and rectangular velocity tables.'

import math
import clr

from pyrevit import forms
from pyrevit import script

try:
    clr.AddReference('System.Data')
except Exception:
    pass

try:
    from System.Data import DataTable
except Exception:
    DataTable = None

# WPF colour imports for velocity-band fills
_WPF_COLOR_OK = False
try:
    from System.Windows.Media import SolidColorBrush, Color
    from System.Windows.Data import IValueConverter, Binding
    from System.Windows import FrameworkElementFactory, DataTemplate, Thickness
    from System.Windows.Controls import TextBlock, DataGridTemplateColumn, Border, DataGridLength
    import System.Windows
    _WPF_COLOR_OK = True
except Exception:
    pass

# ---------------------------------------------------------------------------
# Velocity → colour mapping (exact Excel conditional-formatting colours)
#   < 2.0  m/s  →  #CCFFFF  light cyan   (indexed 41)
#   2.0–3.0 m/s  →  #FFFFCC  cream yellow (indexed 26)
#   3.0–4.0 m/s  →  #FFCC00  amber/gold   (indexed 51)
#   4.0–5.0 m/s  →  #FF6600  orange       (indexed 53)
#   > 5.0  m/s  →  #FF4444  red          (theme 5 tint 0.4 approx.)
# ---------------------------------------------------------------------------
_VEL_BANDS = [
    (2.0,  (0xCC, 0xFF, 0xFF)),
    (3.0,  (0xFF, 0xFF, 0xCC)),
    (4.0,  (0xFF, 0xCC, 0x00)),
    (5.0,  (0xFF, 0x66, 0x00)),
    (None, (0xFF, 0x44, 0x44)),
]


def _vel_bg_brush(v):
    """Return a SolidColorBrush for velocity *v* (m/s), or None if unavailable."""
    if not _WPF_COLOR_OK:
        return None
    for threshold, rgb in _VEL_BANDS:
        if threshold is None or v < threshold:
            return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    return None


if _WPF_COLOR_OK:
    class VelocityColorConverter(IValueConverter):
        """IValueConverter: velocity string → SolidColorBrush (for SquareGrid cells)."""

        def Convert(self, value, target_type, parameter, culture):
            try:
                v = float(str(value).strip())
                if v <= 0:
                    raise ValueError('non-positive')
            except Exception:
                return SolidColorBrush(Color.FromRgb(0xF9, 0xFA, 0xFB))
            return _vel_bg_brush(v) or SolidColorBrush(Color.FromRgb(0xF9, 0xFA, 0xFB))

        def ConvertBack(self, value, target_type, parameter, culture):
            return None


ROUND_DIAMETERS_MM = [80, 90, 100, 112, 125, 150, 160, 180, 200, 225, 250, 280, 315, 400, 500]
SQUARE_SIZES_MM = list(range(100, 3001, 50))
DISPLAY_WIDTHS_MM = list(range(100, 1500, 50))


def _to_float(value, default_value):
    try:
        return float(str(value).strip())
    except Exception:
        return float(default_value)


def _round_to_nearest_50(value):
    return int(round(value / 50.0) * 50)


def _velocity_from_round(airflow_lps, diameter_mm):
    area = math.pi * ((diameter_mm / 1000.0) ** 2) / 4.0
    if area <= 0:
        return 0.0
    return (airflow_lps / 1000.0) / area


def _velocity_from_rect(airflow_lps, width_mm, height_mm):
    area = (width_mm / 1000.0) * (height_mm / 1000.0)
    if area <= 0:
        return 0.0
    return (airflow_lps / 1000.0) / area


def _round_diameter_for_velocity(airflow_lps, velocity_mps):
    if airflow_lps <= 0 or velocity_mps <= 0:
        return 0.0
    flow_m3s = airflow_lps / 1000.0
    diameter_m = math.sqrt((4.0 * flow_m3s) / (math.pi * velocity_mps))
    return diameter_m * 1000.0


def _suggest_height(airflow_lps, velocity_mps, width_mm):
    if airflow_lps <= 0 or velocity_mps <= 0 or width_mm <= 0:
        return 0
    flow_m3s = airflow_lps / 1000.0
    width_m = width_mm / 1000.0
    h_mm = (flow_m3s / (velocity_mps * width_m)) * 1000.0
    h_mm = max(100.0, h_mm)
    return _round_to_nearest_50(h_mm)


class DuctulatorWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)
        if _WPF_COLOR_OK:
            self._vel_converter = VelocityColorConverter()
            # Wire AutoGeneratingColumn BEFORE ItemsSource is set in _recalculate
            self.SquareGrid.AutoGeneratingColumn += self.square_auto_col
        self._seed_legend()
        self._recalculate()

    # ------------------------------------------------------------------
    # AutoGeneratingColumn handler – replaces each numeric column in the
    # square-duct DataGrid with a colour-coded DataGridTemplateColumn.
    # ------------------------------------------------------------------
    def square_auto_col(self, sender, args):
        if not _WPF_COLOR_OK:
            return
        try:
            col_name = str(args.PropertyName)
            if col_name == 'h x w':
                return

            template_col            = DataGridTemplateColumn()
            template_col.Header     = col_name
            template_col.Width      = DataGridLength(65)
            template_col.IsReadOnly = True

            border_fac = FrameworkElementFactory(Border)
            bg_bind    = Binding(col_name)
            bg_bind.Converter = self._vel_converter
            border_fac.SetBinding(Border.BackgroundProperty, bg_bind)

            text_fac = FrameworkElementFactory(TextBlock)
            txt_bind = Binding(col_name)
            text_fac.SetBinding(TextBlock.TextProperty, txt_bind)
            text_fac.SetValue(
                TextBlock.TextAlignmentProperty,
                System.Windows.TextAlignment.Center)
            text_fac.SetValue(
                TextBlock.PaddingProperty,
                Thickness(4.0, 2.0, 4.0, 2.0))
            text_fac.SetValue(
                TextBlock.ForegroundProperty,
                SolidColorBrush(Color.FromRgb(0x1F, 0x29, 0x37)))

            border_fac.AppendChild(text_fac)

            cell_template             = DataTemplate()
            cell_template.VisualTree  = border_fac
            template_col.CellTemplate = cell_template

            args.Column = template_col
        except Exception:
            # Never let an exception escape into WPF layout – that crashes Revit
            pass

    def _seed_legend(self):
        bands = [
            ('< 2.0',     'Very Low',  1.0),
            ('2.0 - 3.0', 'Low',       2.5),
            ('3.0 - 4.0', 'Preferred', 3.5),
            ('4.0 - 5.0', 'High',      4.5),
            ('> 5.0',     'Very High', 5.5),
        ]
        if DataTable is not None:
            table = DataTable()
            table.Columns.Add('Range')
            table.Columns.Add('Status')
            if _WPF_COLOR_OK:
                table.Columns.Add('BgColor')
            for rng, status, v in bands:
                row = table.NewRow()
                row['Range']  = rng
                row['Status'] = status
                if _WPF_COLOR_OK:
                    row['BgColor'] = _vel_bg_brush(v)
                table.Rows.Add(row)
            self.LegendGrid.ItemsSource = table.DefaultView
        else:
            # Fallback – plain list (no binding, values won't show colours)
            rows = []
            for rng, status, v in bands:
                rows.append({'Range': rng, 'Status': status})
            self.LegendGrid.ItemsSource = rows

    def _set_summary_text(self, airflow_lps, velocity_mps, width_mm):
        self.TxtFlowrate.Text = '{0:.1f}'.format(airflow_lps * 3.6)
        round_dia = _round_diameter_for_velocity(airflow_lps, velocity_mps)
        self.TxtRoundDia.Text = '{0:.0f}'.format(round_dia)

        height_mm = _suggest_height(airflow_lps, velocity_mps, width_mm)
        self.TxtRect.Text = '{0:.0f} x {1}'.format(width_mm, height_mm)

        self.TxtHeader.Text = (
            'Airflow {0:.1f} L/s  |  Velocity {1:.2f} m/s  |  Known width {2:.0f} mm'
            .format(airflow_lps, velocity_mps, width_mm)
        )

    def _set_round_table(self, airflow_lps):
        if DataTable is not None:
            table = DataTable()
            table.Columns.Add('Diameter')
            table.Columns.Add('Velocity')
            if _WPF_COLOR_OK:
                table.Columns.Add('BgColor')
            for d_mm in ROUND_DIAMETERS_MM:
                v   = _velocity_from_round(airflow_lps, d_mm)
                row = table.NewRow()
                row['Diameter'] = str(d_mm)
                row['Velocity'] = '{0:.2f}'.format(v)
                if _WPF_COLOR_OK:
                    row['BgColor'] = _vel_bg_brush(v)
                table.Rows.Add(row)
            self.RoundGrid.ItemsSource = table.DefaultView
        else:
            rows = []
            for d_mm in ROUND_DIAMETERS_MM:
                v = _velocity_from_round(airflow_lps, d_mm)
                rows.append({'Diameter': str(d_mm), 'Velocity': '{0:.2f}'.format(v)})
            self.RoundGrid.ItemsSource = rows

    def _set_square_table(self, airflow_lps):
        if DataTable is not None:
            table = DataTable()
            table.Columns.Add('h x w')

            for width_mm in DISPLAY_WIDTHS_MM:
                table.Columns.Add(str(width_mm))

            for height_mm in SQUARE_SIZES_MM:
                row = table.NewRow()
                row['h x w'] = str(height_mm)
                for width_mm in DISPLAY_WIDTHS_MM:
                    v = _velocity_from_rect(airflow_lps, width_mm, height_mm)
                    row[str(width_mm)] = '{0:.2f}'.format(v)
                table.Rows.Add(row)

            self.SquareGrid.ItemsSource = table.DefaultView
            return

        rows = []
        for height_mm in SQUARE_SIZES_MM:
            row = {'h x w': str(height_mm)}
            for width_mm in DISPLAY_WIDTHS_MM:
                v = _velocity_from_rect(airflow_lps, width_mm, height_mm)
                row[str(width_mm)] = '{0:.2f}'.format(v)
            rows.append(row)
        self.SquareGrid.ItemsSource = rows

    def _recalculate(self):
        airflow_lps = _to_float(self.InpAirflow.Text, 240.0)
        velocity_mps = _to_float(self.InpVelocity.Text, 4.0)
        width_mm = _to_float(self.InpWidth.Text, 250.0)

        if airflow_lps <= 0:
            airflow_lps = 240.0
        if velocity_mps <= 0:
            velocity_mps = 4.0
        if width_mm <= 0:
            width_mm = 250.0

        self.InpAirflow.Text = '{0:g}'.format(airflow_lps)
        self.InpVelocity.Text = '{0:g}'.format(velocity_mps)
        self.InpWidth.Text = '{0:g}'.format(width_mm)

        self._set_summary_text(airflow_lps, velocity_mps, width_mm)
        self._set_round_table(airflow_lps)
        self._set_square_table(airflow_lps)

    def calculate_click(self, sender, args):
        self._recalculate()

    def reset_click(self, sender, args):
        self.InpAirflow.Text = '240'
        self.InpVelocity.Text = '4'
        self.InpWidth.Text = '250'
        self._recalculate()

    def help_click(self, sender, args):
        forms.alert(
            'How to use Ductulator:\n\n'
            '1) Enter Airflow in L/s.\n'
            '2) Enter a target velocity in m/s.\n'
            '3) Enter known duct width in mm.\n'
            '4) Click Calculate.\n\n'
            'The tool updates:\n'
            '- Flowrate in m3/h\n'
            '- Round equivalent diameter\n'
            '- Suggested rectangular size\n'
            '- Round and square velocity tables for quick checks.\n',
            title='Ductulator Instructions'
        )


def main():
    xaml_path = script.get_bundle_file('Ductulator.xaml')
    window = DuctulatorWindow(xaml_path)
    window.ShowDialog()


if __name__ == '__main__':
    main()
