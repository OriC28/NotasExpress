from PyQt6 import QtCore
# NotasExpress
# Copyright (C) 2025 Escuela Técnica Industrial Nacional Capitán Giovanni Ferrareis 
# Licenciado bajo la GNU GPLv3. Ver <https://www.gnu.org/licenses/>.

class InitAppController:
    """
    Controlador para la inicialización de la aplicación.

    Este controlador se encarga de configurar la interfaz gráfica y conectar las señales
    y slots necesarios para el funcionamiento de la aplicación. También establece las
    conexiones entre las señales de progreso y los métodos correspondientes.

    Métodos:
        init_app: Configura la interfaz gráfica y conecta las señales y slots.
    """

    def init_app(self):
        """
        Inicializa la aplicación y configura la interfaz gráfica.

        Este método realiza las siguientes acciones:
        1. Establece el logo de la aplicación en la vista principal.
        2. Conecta los botones de la interfaz gráfica con sus respectivos métodos:
           - `SelectButton` con `select_excel_file`.
           - `SaveButton` con `select_path_to_save`.
           - `GenerateButton` con `generate_notes`.
        3. Conecta el cambio de índice en el combo box `CbYearSection` con el método `fill_table`.
        4. Muestra la vista principal de la aplicación.
        5. Conecta las señales de progreso con los métodos correspondientes:
           - `convertion_started` con `show_loading_dialog`.
           - `convertion_finished` con `on_conversion_finished`.
           - `progress_updated` con `loading_dialog.update_progress`.
        """
        self.view.set_logo()
        self.view.SelectButton.clicked.connect(self.select_excel_file)
        self.view.SaveButton.clicked.connect(self.select_path_to_save)
        self.view.GenerateButton.clicked.connect(self.generate_notes)
        self.view.CbYearSection.currentIndexChanged.connect(self.fill_table)
        QtCore.QCoreApplication.instance().aboutToQuit.connect(self.view.closeEvent)
        QtCore.QCoreApplication.instance().aboutToQuit.connect(self.loading_dialog.closeEvent)
        self.view.show()

        self.convertion_started.connect(self.show_loading_dialog)
        self.convertion_finished.connect(self.on_conversion_finished)
        self.progress_updated.connect(self.loading_dialog.update_progress)