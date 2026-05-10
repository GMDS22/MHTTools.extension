# -*- coding: utf-8 -*-
from __future__ import division, print_function

__title__ = 'Ductulator'
__author__ = 'GMoreno'
__doc__ = 'ENVAR duct sizing reference with round and rectangular velocity tables.'
__persistentengine__ = True

import math
import clr

from pyrevit import forms
from pyrevit import script

# WPF colour imports for velocity-band fills
_WPF_COLOR_OK = False
try:
    from System.Windows.Media import SolidColorBrush, Color
    from System.Windows.Data import Binding
    from System.Windows import FrameworkElementFactory, DataTemplate, Thickness
    from System.Windows.Controls import TextBlock, Border, DataGridLength, DataGridTemplateColumn, DataGridTextColumn
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


def _vel_bg_hex(v):
    """Return a WPF-compatible hex colour string for velocity *v* (m/s)."""
    for threshold, rgb in _VEL_BANDS:
        if threshold is None or v < threshold:
            return '#{0:02X}{1:02X}{2:02X}'.format(rgb[0], rgb[1], rgb[2])
    return '#FFFFFF'


def _vel_bg_value(v):
    brush = _vel_bg_brush(v)
    if brush is not None:
        return brush
    return _vel_bg_hex(v)


class _GridRow(object):
    def __init__(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self, key, value)


if _WPF_COLOR_OK:
    pass


ROUND_DIAMETERS_MM = [80, 90, 100, 112, 125, 150, 160, 180, 200, 225, 250, 280, 315, 400, 500]
SQUARE_SIZES_MM = list(range(100, 3001, 50))
DISPLAY_WIDTHS_MM = list(range(100, 1500, 50))


# Keep a strong reference for modeless window lifetime across script calls.
_DUCTULATOR_WINDOW = None


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


def _pa_per_m_rect(airflow_lps, width_mm, height_mm,
                   roughness_m=0.0001, rho=1.2, nu=1.5e-5):
    """Darcy-Weisbach friction pressure gradient (Pa/m) for a rectangular duct.

    Args:
        airflow_lps  : airflow in L/s
        width_mm     : duct width in mm  (cross-section dimension)
        height_mm    : duct height in mm (cross-section dimension, "Duct Length")
        roughness_m  : absolute roughness in m (0.1 mm default for sheet metal)
        rho          : air density kg/m³ (1.2 at ~20°C sea level)
        nu           : kinematic viscosity m²/s (1.5e-5 at ~20°C)

    Returns:
        pressure gradient in Pa/m
    """
    if width_mm <= 0 or height_mm <= 0 or airflow_lps <= 0:
        return 0.0
    W = width_mm / 1000.0
    H = height_mm / 1000.0
    A = W * H
    D_h = 4.0 * A / (2.0 * (W + H))          # hydraulic diameter
    V = (airflow_lps / 1000.0) / A            # mean velocity m/s
    Re = V * D_h / nu
    if Re < 1.0:
        return 0.0
    # Haaland (1983) explicit approximation of Colebrook-White — valid Re > 3000
    # Falls back gracefully for laminar-transitional range
    inner = (roughness_m / D_h / 3.7) ** 1.11 + 6.9 / Re
    if inner <= 0:
        return 0.0
    f = (1.0 / (-1.8 * math.log10(inner))) ** 2
    return f * rho * V ** 2 / (2.0 * D_h)


class DuctulatorWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)
        self._build_square_grid_columns()
        self._seed_legend()
        self._recalculate()

    def _build_square_grid_columns(self):
        if not _WPF_COLOR_OK:
            return
        try:
            self.SquareGrid.Columns.Clear()

            h_col = DataGridTextColumn()
            h_col.Header = 'H x W'
            h_col.Binding = Binding('HxW')
            h_col.Width = DataGridLength(70)
            self.SquareGrid.Columns.Add(h_col)

            for width_mm in DISPLAY_WIDTHS_MM:
                w_key = 'W{0}'.format(width_mm)
                c_key = 'C{0}'.format(width_mm)

                txt = FrameworkElementFactory(TextBlock)
                txt.SetBinding(TextBlock.TextProperty, Binding(w_key))
                txt.SetBinding(TextBlock.BackgroundProperty, Binding(c_key))
                txt.SetValue(TextBlock.TextAlignmentProperty, System.Windows.TextAlignment.Center)
                txt.SetValue(TextBlock.PaddingProperty, Thickness(4, 2, 4, 2))
                txt.SetValue(TextBlock.ForegroundProperty, SolidColorBrush(Color.FromRgb(0x00, 0x00, 0x00)))

                cell_border = FrameworkElementFactory(Border)
                cell_border.SetValue(Border.PaddingProperty, Thickness(0))
                cell_border.AppendChild(txt)

                template = DataTemplate()
                template.VisualTree = cell_border

                col = DataGridTemplateColumn()
                col.Header = str(width_mm)
                col.Width = DataGridLength(62)
                col.CellTemplate = template
                self.SquareGrid.Columns.Add(col)
        except Exception:
            # Keep the tool alive even if custom columns fail in a given host.
            pass

    def _seed_legend(self):
        bands = [
            ('< 2.0',     'Very Low',  1.0),
            ('2.0 - 3.0', 'Low',       2.5),
            ('3.0 - 4.0', 'Preferred', 3.5),
            ('4.0 - 5.0', 'High',      4.5),
            ('> 5.0',     'Very High', 5.5),
        ]
        rows = []
        for rng, status, v in bands:
            rows.append(_GridRow(Range=rng, Status=status, BgColor=_vel_bg_value(v)))
        self.LegendGrid.ItemsSource = rows

    def _set_summary_text(self, airflow_lps, velocity_mps, width_mm, duct_length_mm):
        self.TxtFlowrate.Text = '{0:.1f}'.format(airflow_lps * 3.6)
        round_dia = _round_diameter_for_velocity(airflow_lps, velocity_mps)
        self.TxtRoundDia.Text = '{0:.0f}'.format(round_dia)

        # Use explicit duct length (height) if provided, otherwise suggest from velocity+width.
        if duct_length_mm > 0:
            height_mm = duct_length_mm
        else:
            height_mm = _suggest_height(airflow_lps, velocity_mps, width_mm)
        # Always present rectangular size as W x H where W is the larger dimension.
        width_out = max(width_mm, height_mm)
        height_out = min(width_mm, height_mm)
        self.TxtRect.Text = '{0} x {1}'.format(int(round(width_out)), int(round(height_out)))

        # Pa/m Loss — only meaningful when both cross-section dimensions are known.
        if duct_length_mm > 0:
            pa_m = _pa_per_m_rect(airflow_lps, width_mm, duct_length_mm)
            self.TxtPaPerM.Text = '{0:.2f}'.format(pa_m)
        else:
            self.TxtPaPerM.Text = '—'

        self.TxtHeader.Text = (
            'Airflow {0:.1f} L/s  |  Velocity {1:.2f} m/s  |  Width {2:.0f} mm  |  Duct Length {3}'
            .format(
                airflow_lps,
                velocity_mps,
                width_mm,
                '{0:.0f} mm'.format(duct_length_mm) if duct_length_mm > 0 else 'not set',
            )
        )

    def _set_round_table(self, airflow_lps):
        rows = []
        for d_mm in ROUND_DIAMETERS_MM:
            v = _velocity_from_round(airflow_lps, d_mm)
            rows.append(
                _GridRow(
                    Diameter=str(d_mm),
                    Velocity='{0:.2f}'.format(v),
                    BgColor=_vel_bg_value(v)
                )
            )
        self.RoundGrid.ItemsSource = rows

    def _set_square_table(self, airflow_lps):
        # Only clear ItemsSource on refresh; columns are created once at init.
        try:
            self.SquareGrid.ItemsSource = None
        except Exception:
            pass

        rows = []
        for height_mm in SQUARE_SIZES_MM:
            row_data = {'HxW': str(height_mm)}
            for width_mm in DISPLAY_WIDTHS_MM:
                v = _velocity_from_rect(airflow_lps, width_mm, height_mm)
                row_data['W{0}'.format(width_mm)] = '{0:.2f}'.format(v)
                row_data['C{0}'.format(width_mm)] = _vel_bg_value(v)
            rows.append(_GridRow(**row_data))
        self.SquareGrid.ItemsSource = rows

    def _recalculate(self):
        airflow_lps = _to_float(self.InpAirflow.Text, 240.0)
        velocity_mps = _to_float(self.InpVelocity.Text, 4.0)
        width_mm = _to_float(self.InpWidth.Text, 250.0)
        # Duct Length = second cross-section dimension (height). 0 means 'not set'.
        duct_length_mm = _to_float(self.InpDuctLength.Text, 0.0)
        if duct_length_mm < 0:
            duct_length_mm = 0.0

        if airflow_lps <= 0:
            airflow_lps = 240.0
        if velocity_mps <= 0:
            velocity_mps = 4.0
        if width_mm <= 0:
            width_mm = 250.0

        self.InpAirflow.Text = '{0:g}'.format(airflow_lps)
        self.InpVelocity.Text = '{0:g}'.format(velocity_mps)
        self.InpWidth.Text = '{0:g}'.format(width_mm)
        # Leave InpDuctLength blank when zero so the placeholder hint shows.
        if duct_length_mm > 0:
            self.InpDuctLength.Text = '{0:g}'.format(duct_length_mm)

        self._set_summary_text(airflow_lps, velocity_mps, width_mm, duct_length_mm)
        self._set_round_table(airflow_lps)
        self._set_square_table(airflow_lps)

    def calculate_click(self, sender, args):
        try:
            self._recalculate()
        except Exception as ex:
            forms.alert(
                'Calculation failed and was safely stopped.\n\n{0}'.format(str(ex)),
                title='MHT Ductolator by GM'
            )

    def reset_click(self, sender, args):
        try:
            self.InpAirflow.Text = '240'
            self.InpVelocity.Text = '4'
            self.InpWidth.Text = '250'
            self.InpDuctLength.Text = ''
            self._recalculate()
        except Exception as ex:
            forms.alert(
                'Reset failed and was safely stopped.\n\n{0}'.format(str(ex)),
                title='MHT Ductolator by GM'
            )

    def help_click(self, sender, args):
        forms.alert(
            'How to use Ductulator:\n\n'
            '1) Enter Airflow in L/s.\n'
            '2) Enter a target velocity in m/s.\n'
            '3) Enter known Duct Width in mm.\n'
            '4) (Optional) Enter Duct Length in mm \u2014 the second cross-section\n'
            '   dimension (height of the duct face). When provided, Pa/m Loss\n'
            '   is calculated using Darcy-Weisbach with the Haaland friction\n'
            '   factor (sheet-metal roughness 0.1 mm, air at 20\u00b0C).\n'
            '5) Click Calculate.\n\n'
            'The tool updates:\n'
            '- Flowrate in m\u00b3/h\n'
            '- Round equivalent diameter\n'
            '- Suggested rectangular size\n'
            '- Pa/m Loss (friction pressure gradient) \u2014 requires both Duct\n'
            '  Width and Duct Length to be set\n'
            '- Round and rectangular velocity tables for quick checks.\n',
            title='Ductulator Instructions'
        )


def main():
    global _DUCTULATOR_WINDOW
    xaml_path = script.get_bundle_file('Ductulator.xaml')
    if _DUCTULATOR_WINDOW is not None:
        try:
            _DUCTULATOR_WINDOW.Close()
        except Exception:
            pass
    _DUCTULATOR_WINDOW = DuctulatorWindow(xaml_path)
    _DUCTULATOR_WINDOW.Show()
    try:
        _DUCTULATOR_WINDOW.Activate()
    except Exception:
        pass


if __name__ == '__main__':
    main()
