import unittest
from unittest.mock import patch, MagicMock
from scripts.interface_v3 import GLPIApp
import pandas as pd

class TestGLPIApp(unittest.TestCase):

    @patch('interface_v3.simpledialog.askstring')
    @patch('interface_v3.GLPIApp.escanear_qr_con_celular')
    @patch('interface_v3.GLPIApp.verificar_existencia_en_excel')
    @patch('interface_v3.GLPIApp.agregar_a_excel')
    def test_manejar_qr_laptop_register(self, mock_agregar_a_excel, mock_verificar_existencia_en_excel, mock_escanear_qr_con_celular, mock_askstring):
        # Mock the user inputs and QR scan result
        mock_askstring.side_effect = ["Dell", "escanear", "sí", "Office"]
        mock_escanear_qr_con_celular.return_value = "8B9X1R3"
        mock_verificar_existencia_en_excel.return_value = False

        app = GLPIApp(None)
        app.manejar_qr_laptop("Register")

        # Verify that agregar_a_excel was called with the correct data
        expected_asset_data = {
            "Asset Type": "Computer",
            "Status": "Stocked",
            "User": None,
            "Name": "None-Latitude",
            "Computer Types": "Laptop",
            "Location": "Office",
            "Manufacturer": "Dell inc.",
            "Model": "Latitude",
            "Serial Number": "8B9X1R3",
            "Comments": "Check",
        }
        mock_agregar_a_excel.assert_called_once_with(expected_asset_data)

    @patch('interface_v3.simpledialog.askstring')
    @patch('interface_v3.messagebox.showinfo')
    @patch('interface_v3.messagebox.showerror')
    def test_entregar_laptop(self, mock_showerror, mock_showinfo, mock_askstring):
        # Mock the user inputs
        mock_askstring.side_effect = ["Dell", "manual", "8B9X1R3", "sí", "John Doe"]

        app = GLPIApp(None)
        app.manejar_qr_laptop = MagicMock(return_value=("8B9X1R3", "Dell"))
        app.obtener_token_sesion = MagicMock(return_value="mock_token")
        app.obtener_asset_id_por_serial = MagicMock(return_value="mock_asset_id")
        app.actualizar_asset_glpi = MagicMock()

        # Mock the Excel data
        app.df = pd.DataFrame({
            "Serial Number": ["8B9X1R3"],
            "Manufacturer": ["Dell"],
            "User": [""],
            "Name": ["Unknown"]
        })

        app.entregar_laptop()

        # Verify that the asset was updated correctly
        app.actualizar_asset_glpi.assert_called_once_with("mock_token", "mock_asset_id", {
            "Serial Number": "8B9X1R3",
            "Manufacturer": "Dell",
            "User": "John Doe",
            "Name": "John Doe-Latitude"
        })

    @patch('interface_v3.simpledialog.askstring')
    @patch('interface_v3.GLPIApp.escanear_qr_con_celular')
    @patch('interface_v3.GLPIApp.verificar_existencia_en_excel')
    @patch('interface_v3.GLPIApp.agregar_a_excel')
    def test_manejar_qr_monitor(self, mock_agregar_a_excel, mock_verificar_existencia_en_excel, mock_escanear_qr_con_celular, mock_askstring):
        # Mock the user inputs and QR scan result
        mock_askstring.side_effect = ["Office"]
        mock_escanear_qr_con_celular.return_value = "CN0V7X9J129025AN2AX1"
        mock_verificar_existencia_en_excel.return_value = False

        app = GLPIApp(None)
        app.manejar_qr_monitor()

        # Verify that agregar_a_excel was called with the correct data
        expected_asset_data = {
            "Asset Type": "Monitor",
            "Status": "Stocked",
            "User": None,
            "Name": "None-CN0V7X9J129025AN2AX1",
            "Location": "Office",
            "Manufacturer": "Dell Inc.",
            "Model": None,
            "Serial Number": "CN0V7X9J129025AN2AX1",
            "Comments": "Check",
        }
        mock_agregar_a_excel.assert_called_once_with(expected_asset_data)

    @patch('interface_v3.simpledialog.askstring')
    @patch('interface_v3.messagebox.showinfo')
    @patch('interface_v3.messagebox.showerror')
    def test_entregar_monitor(self, mock_showerror, mock_showinfo, mock_askstring):
        # Mock the user inputs
        mock_askstring.side_effect = ["escanear", "CN0V7X9J129025AN2AX1", "sí", "John Doe"]

        app = GLPIApp(None)
        app.manejar_qr_monitor = MagicMock(return_value=("CN0V7X9J129025AN2AX1", "Dell"))
        app.obtener_token_sesion = MagicMock(return_value="mock_token")
        app.obtener_asset_id_por_serial_monitor = MagicMock(return_value="mock_asset_id")
        app.actualizar_asset_glpi_monitor = MagicMock()

        # Mock the Excel data
        app.df = pd.DataFrame({
            "Serial Number": ["CN0V7X9J129025AN2AX1"],
            "Manufacturer": ["Dell"],
            "User": [""],
            "Name": ["Unknown"]
        })

        app.entregar_monitor()

        # Verify that the asset was updated correctly
        app.actualizar_asset_glpi_monitor.assert_called_once_with("mock_token", "mock_asset_id", {
            "Serial Number": "CN0V7X9J129025AN2AX1",
            "Manufacturer": "Dell",
            "User": "John Doe",
            "Name": "John Doe-DellMonitor"
        })

if __name__ == '__main__':
    unittest.main()