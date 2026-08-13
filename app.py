from gui import MainWindow
from controller import Controller

class App(MainWindow):
    def __init__(self):
        super().__init__()
        # Создаём контроллер и связываем с GUI
        self.controller = Controller(self, self.current_folder)
        # Даём GUI ссылку на контроллер
        self.set_controller(self.controller)

if __name__ == "__main__":
    app = App()
    app.mainloop()