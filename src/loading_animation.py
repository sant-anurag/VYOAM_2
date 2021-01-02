import pyglet
from pyglet import *


# Class definition for member donation statement
class LoadingAnimation:
    def __init__(self):
        print("constructor called for LoadingAnimation ")
        self.window = 0
        self.load_animation()

    def load_animation(self):
        vidPath = "..\\Images\\Logos\\VYOAM_loading.mp4"
        self.window = pyglet.window.Window(style=pyglet.window.Window.WINDOW_STYLE_BORDERLESS, vsync=False)
        x, y = self.window.get_location()
        self.window.set_location(x + 400, y + 200)
        self.window.set_size(950, 500)
        self.window.set_caption(
            '                                                                                                                                Vihangam Yoga Karnataka')
        self.player = pyglet.media.Player()
        source = pyglet.media.StreamingSource()
        MediaLoad = pyglet.media.load(vidPath)

        self.player.queue(MediaLoad)
        self.player.play()
        pyglet.clock.schedule_once(self.exit_callback, 6)
        pyglet.app.run()

    def exit_callback(self, m):
        print("exit_callback", m)
        self.window.close()


    def on_draw(self):
        if self.player.source and self.player.source.video_format:
            self.player.get_texture().blit(0, 0)
