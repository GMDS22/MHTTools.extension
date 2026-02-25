"""Reload pyRevit into new session.
Shows a small animated loading window while reloading.
"""
# -*- coding: utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except

from pyrevit import script
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo


logger = script.get_logger()
results = script.get_results()

w = None
try:
    xamlfile = script.get_bundle_file('ReloadingWindow.xaml')

    # Try to load the animated GIF
    giffile = None
    for gif_name in ['loading_v2.gif', 'loading.gif']:
        try:
            giffile = script.get_bundle_file(gif_name)
            if giffile:
                logger.debug('Using GIF: {}'.format(gif_name))
                break
        except Exception as ex:
            logger.debug('Could not load {}: {}'.format(gif_name, ex))
    
    if not giffile:
        logger.warning('No GIF file found for loading animation')

    import clr
    try:
        clr.AddReference('WindowsFormsIntegration')
    except Exception:
        pass
    try:
        clr.AddReference('System.Windows.Forms')
    except Exception:
        pass
    try:
        clr.AddReference('System.Drawing')
    except Exception:
        pass

    import wpf
    from System import Windows
    from System.Windows.Threading import DispatcherTimer
    from System import TimeSpan

    from System.Windows.Forms import PictureBox
    from System.Windows.Forms import PictureBoxSizeMode
    from System.Drawing import Image
    from System.Drawing.Imaging import FrameDimension

    def _close_existing_reload_windows():
        try:
            for win in Windows.Application.Current.Windows:
                try:
                    if hasattr(win, 'FindName') and win.FindName('GifHost') is not None:
                        try:
                            win.Close()
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

    class ReloadingWindow(Windows.Window):
        def __init__(self):
            wpf.LoadComponent(self, xamlfile)

            self._pb = None
            self._gif_image = None
            self._gif_dim = None
            self._gif_frame_count = 0
            self._gif_frame_index = 0
            self._anim_timer = None

            try:
                self._pb = PictureBox()
                self._pb.SizeMode = PictureBoxSizeMode.Zoom

                if giffile:
                    logger.debug('Loading GIF from: {}'.format(giffile))
                    self._gif_image = Image.FromFile(giffile)
                    logger.debug('GIF loaded successfully')
                else:
                    logger.warning('No GIF file provided')

                if self._gif_image:
                    try:
                        self._gif_dim = FrameDimension(self._gif_image.FrameDimensionsList[0])
                        self._gif_frame_count = self._gif_image.GetFrameCount(self._gif_dim)
                        logger.debug('GIF has {} frames'.format(self._gif_frame_count))
                    except Exception as ex:
                        logger.debug('Error getting frame info: {}'.format(ex))
                        self._gif_dim = None
                        self._gif_frame_count = 1

                self._pb.Image = self._gif_image
                self.GifHost.Child = self._pb

                try:
                    if self._gif_dim is not None and self._gif_frame_count > 1:
                        self._anim_timer = DispatcherTimer()
                        self._anim_timer.Interval = TimeSpan.FromMilliseconds(100)
                        self._anim_timer.Tick += self._on_anim_tick
                        self._anim_timer.Start()
                        logger.debug('Animation timer started for {} frames'.format(self._gif_frame_count))
                except Exception as ex:
                    logger.debug('Could not start animation timer: {}'.format(ex))
                    self._anim_timer = None
            except Exception as e:
                logger.debug('Could not load animated gif: {}'.format(e))

            try:
                self._timeout_timer = DispatcherTimer()
                self._timeout_timer.Interval = TimeSpan.FromSeconds(20)
                self._timeout_timer.Tick += self._on_timeout
                self._timeout_timer.Start()
            except Exception:
                self._timeout_timer = None

            try:
                self.Closed += self._on_closed
            except Exception:
                pass

        def _on_anim_tick(self, sender, e):
            try:
                if self._gif_image is None or self._gif_dim is None or self._pb is None:
                    return
                self._gif_frame_index = (self._gif_frame_index + 1) % self._gif_frame_count
                self._gif_image.SelectActiveFrame(self._gif_dim, self._gif_frame_index)
                try:
                    self._pb.Invalidate()
                    self._pb.Refresh()
                except Exception:
                    pass
            except Exception:
                pass

        def _on_timeout(self, sender, e):
            try:
                if self._timeout_timer:
                    self._timeout_timer.Stop()
            except Exception:
                pass
            try:
                self.Close()
            except Exception:
                pass

        def _on_closed(self, sender, e):
            try:
                if self._anim_timer:
                    self._anim_timer.Stop()
            except Exception:
                pass
            try:
                if self._timeout_timer:
                    self._timeout_timer.Stop()
            except Exception:
                pass
            try:
                if self._gif_image is not None:
                    self._gif_image.Dispose()
            except Exception:
                pass

    try:
        _close_existing_reload_windows()
    except Exception as ex:
        logger.debug('Error closing existing windows: {}'.format(ex))

    try:
        w = ReloadingWindow()
        w.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
        w.Show()
        try:
            w.Activate()
            w.Topmost = True
        except Exception as ex:
            logger.debug('Could not activate window: {}'.format(ex))
    except Exception as e:
        logger.warning('Could not display reloading window: {}'.format(e))
        import traceback
        logger.debug(traceback.format_exc())
except Exception as e:
    logger.warning('Could not prepare reloading window: {}'.format(e))
    import traceback
    logger.debug(traceback.format_exc())

logger.info('RELOADING CHANGES')
try:
    sessionmgr.reload_pyrevit()
finally:
    try:
        if w is not None:
            try:
                w.Close()
            except Exception as e:
                logger.debug('Could not close reloading window: {}'.format(e))
    except Exception as e:
        logger.debug('Error while closing reloading window: {}'.format(e))

try:
    results.newsession = sessioninfo.get_session_uuid()
except Exception as e:
    logger.debug('Could not set newsession: {}'.format(e))
