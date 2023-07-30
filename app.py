import os
import shutil
import subprocess
import requests
import win32con
import win32api
import platform
import zipfile
import pickle
import winsound
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# Obtém o nome do sistema operacional
sistema_operacional = platform.system()

if sistema_operacional != 'Windows':
    os.exit()

GITHUB_REPO = "https://api.github.com/repos/OneDefauter/Cortar-Imagens"
version = "v1.0"

try:
    response = requests.get(f"{GITHUB_REPO}/releases/latest")
    response.raise_for_status()
    latest_version = response.json()["tag_name"]
except requests.exceptions.RequestException as e:
    print(f"Erro ao verificar atualizações: {e}")
    latest_version = None

class ImageCropperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cortar Imagens")
        self.image_folder = None
        self.image_list = []
        self.backup_var = tk.BooleanVar(value=False)
        self.extension_var = tk.StringVar(value=".jpg")
        self.quality_var = tk.IntVar(value=80)
        self.crop_altura_var = tk.IntVar(value=1000)
        self.show_success_msg_var = tk.BooleanVar(value=True)
        self.show_rename_success_msg_var = tk.BooleanVar(value=True)

        if not self.check_imagemagick_installed():
            if messagebox.askyesno("ImageMagick não está instalado", "ImageMagick não foi encontrado. Deseja baixar e instalar agora?"):
                self.download_imagemagick()
                messagebox.showinfo("Instalação Concluído", "O ImageMagick foi instalado com sucesso.")
            self.root.destroy()
            return
            
        # Caminho para o diretório atual (onde o código está sendo executado)
        self.current_dir = os.getcwd()

        # Diretório onde o arquivo settings.pickle será salvo
        self.settings_dir = os.path.join(os.environ.get("HOMEDRIVE"), os.environ.get("HOMEPATH"), "Cortar Imagens")

        if not os.path.exists(self.settings_dir):
            os.mkdir(self.settings_dir)

        # Carregar configurações do usuário
        self.load_settings()

        self.create_widgets()

    def save_settings(self):
        settings = {
            "backup_var": self.backup_var.get(),
            "extension_var": self.extension_var.get(),
            "quality_var": self.quality_var.get(),
            "crop_altura_var": self.crop_altura_var.get(),
            "show_success_msg_var": self.show_success_msg_var.get(),
            "show_rename_success_msg_var": self.show_rename_success_msg_var.get(),
        }

        with open(f"{self.settings_dir}/settings.pickle", "wb") as file:
            pickle.dump(settings, file)

    def load_settings(self):
        try:
            with open(f"{self.settings_dir}/settings.pickle", "rb") as file:
                settings = pickle.load(file)

            self.backup_var.set(settings["backup_var"])
            self.extension_var.set(settings["extension_var"])
            self.quality_var.set(settings["quality_var"])
            self.crop_altura_var.set(settings["crop_altura_var"])
            self.show_success_msg_var.set(settings["show_success_msg_var"])
            self.show_rename_success_msg_var.set(settings["show_rename_success_msg_var"])

        except FileNotFoundError:
            # Usar valores padrão caso o arquivo de configurações não exista
            self.backup_var.set(False)
            self.extension_var.set(".jpg")
            self.quality_var.set(80)
            self.crop_altura_var.set(1000)
            self.show_success_msg_var.set(True)
            self.show_rename_success_msg_var.set(True)
        except (pickle.PickleError, KeyError):
            # Tratar qualquer erro de desserialização ou chave ausente
            print("Erro ao carregar as configurações. Usando valores padrão.")
            self.backup_var.set(False)
            self.extension_var.set(".jpg")
            self.quality_var.set(80)
            self.crop_altura_var.set(1000)
            self.show_success_msg_var.set(True)
            self.show_rename_success_msg_var.set(True)

    def create_widgets(self):
        # Botão para adicionar a pasta de imagens
        tk.Button(self.root, text="Adicionar Pasta", command=self.select_image_folder).grid(row=0, column=0, padx=10, pady=10)

        # Caixa para marcar ou desmarcar a opção de backup
        tk.Checkbutton(self.root, text="Fazer Backup", variable=self.backup_var).grid(row=0, column=1, padx=10, pady=10)

        # Lista com as extensões de saída disponíveis
        tk.Label(self.root, text="Extensão de Saída:").grid(row=1, column=0, padx=10, pady=5)
        extensions = [".png", ".jpg"]
        tk.OptionMenu(self.root, self.extension_var, *extensions).grid(row=1, column=1, padx=10, pady=5)

        # Escala para selecionar o nível da imagem
        tk.Label(self.root, text="Nível de Qualidade:").grid(row=2, column=0, padx=10, pady=5)
        tk.Scale(self.root, from_=1, to=100, variable=self.quality_var, orient=tk.HORIZONTAL).grid(row=2, column=1, padx=10, pady=5)

        # Botão para juntar as imagens
        tk.Button(self.root, text="Cortar Imagens", command=self.crop_images).grid(row=8, column=1, padx=10, pady=10)

        # Botão para renomear os arquivos da pasta
        tk.Button(self.root, text="Renomear", command=self.rename_files).grid(row=8, column=0, padx=10, pady=10)

        # Checkbox para mostrar a mensagem de sucesso após cortar
        tk.Checkbutton(self.root, text="Mostrar mensagem de sucesso após cortar", variable=self.show_success_msg_var).grid(row=5, column=0, columnspan=2)

        # Checkbox para mostrar a mensagem de sucesso após a renomeação
        tk.Checkbutton(self.root, text="Mostrar mensagem de sucesso após renomear", variable=self.show_rename_success_msg_var).grid(row=6, column=0, columnspan=2)

        self.crop_altura_var = tk.StringVar()
        self.crop_altura_var.set("1000")
        tk.Label(self.root, text="Altura do Corte:").grid(row=3, column=0, padx=10, pady=5)
        tk.Entry(self.root, textvariable=self.crop_altura_var).grid(row=3, column=1, padx=10, pady=5)
        
        # Rótulo para mostrar o caminho completo da pasta de imagens
        self.image_folder_label = tk.Label(self.root, text="", wraplength=300, anchor="sw", justify="left")
        self.image_folder_label.grid(row=9, column=0, columnspan=2, padx=10, pady=10)

    def select_image_folder(self):
        self.image_folder = filedialog.askdirectory(title="Selecione a pasta de imagens")
        self.load_image_list()
        
        # Atualizar o rótulo com o caminho completo da pasta de imagens
        self.image_folder_label.config(text=f"Pasta Selecionada:\n{self.image_folder}")

    def load_image_list(self):
        if self.image_folder:
            image_files = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

    def crop_images(self):
        if not self.image_folder:
            messagebox.showerror("Erro", "Selecione uma pasta de imagens primeiro.")
            return

        image_files = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

        input_images = [os.path.join(self.image_folder, image) for image in image_files]
        output_folder = os.path.join(self.image_folder, "temp")
        backup = self.backup_var.get()
        extension = self.extension_var.get()
        quality = self.quality_var.get()
        altura = int(self.crop_altura_var.get())

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        atributos_atuais = win32api.GetFileAttributes(output_folder)
        win32api.SetFileAttributes(output_folder, atributos_atuais | win32con.FILE_ATTRIBUTE_HIDDEN)


        try:
            output_filename = os.path.join(output_folder, f"0{extension}")
            command = ["magick", "convert", "-quality", str(quality), "-crop", f"32000x{altura}"]
            
            command += input_images + [output_filename]
            
            subprocess.run(command, check=True)

            if backup:
                backup_path = os.path.join(self.image_folder, "Backup")
                if not os.path.exists(backup_path):
                    os.makedirs(backup_path)

                for image_file in input_images:
                    shutil.move(image_file, backup_path)
            else:
                for image_file in input_images:
                    source_path = os.path.join(self.image_folder, image_file)
                    os.remove(source_path)

            # Contador para numerar os arquivos
            count = 1

            output_files = sorted([f for f in os.listdir(output_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))], key=lambda x: [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', x)])

            for filename in output_files:
                base, ext = os.path.splitext(filename)
                new_filename = f"{count:02d}{ext}"
                os.rename(os.path.join(output_folder, filename), os.path.join(output_folder, new_filename))
                count += 1

            output_files = [f for f in os.listdir(output_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            for image in output_files:
                output_pathfile = os.path.join(output_folder, image)
                shutil.move(output_pathfile, self.image_folder)
            
            # shutil.move(output_filename, self.image_folder)
            output_folder2 = os.path.join(self.image_folder, "temp")
            os.removedirs(output_folder2)

            # Exibir a mensagem de sucesso apenas se a caixa de seleção estiver marcada
            if self.show_success_msg_var.get():
                messagebox.showinfo("Sucesso", "As imagens foram juntadas com sucesso!")
            else:
                winsound.Beep(1000, 500)  # O primeiro argumento é a frequência em Hz e o segundo é a duração em milissegundos

            self.load_image_list()  # Atualizar a lista de imagens após a junção
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao executar o ImageMagick: \n{e}")
        self.save_settings()

    def rename_files(self):
        if not self.image_folder:
            messagebox.showerror("Erro", "Selecione uma pasta de imagens primeiro.")
            return
        
        backup = self.backup_var.get()

        file_list = [f for f in os.listdir(self.image_folder) if os.path.isfile(os.path.join(self.image_folder, f))]
        file_list.sort()

        if backup:
            backup_path = os.path.join(self.image_folder, "Backup")
            if not os.path.exists(backup_path):
                os.makedirs(backup_path)
            
            for image_file in file_list:
                source_path = os.path.join(self.image_folder, image_file)
                shutil.copy(source_path, backup_path)

        # Contador para numerar os arquivos
        count = 1

        for filename in file_list:
            base, ext = os.path.splitext(filename)
            new_filename = f"{base}__{ext}"
            os.rename(os.path.join(self.image_folder, filename), os.path.join(self.image_folder, new_filename))

        file_list = [f for f in os.listdir(self.image_folder) if os.path.isfile(os.path.join(self.image_folder, f))]
        file_list.sort()

        for filename in file_list:
            base, ext = os.path.splitext(filename)
            new_filename = f"{count:02d}{ext}"
            os.rename(os.path.join(self.image_folder, filename), os.path.join(self.image_folder, new_filename))
            count += 1

        messagebox.showinfo("Sucesso", "Os arquivos foram renomeados com sucesso!") if self.show_rename_success_msg_var.get() else winsound.Beep(1000, 500)  # O primeiro argumento é a frequência em Hz e o segundo é a duração em milissegundos
        self.load_image_list()  # Atualizar a lista de imagens após a renomeação

    def check_imagemagick_installed(self):
        try:
            subprocess.run(["magick", "-version"], capture_output=True, check=True)
            return True
        except subprocess.CalledProcessError:
            try:
                subprocess.run(["convert", "-version"], capture_output=True, check=True)
                return True
            except subprocess.CalledProcessError:
                return False

    def download_newverison(self):
        url = "https://github.com/OneDefauter/Menu_/releases/download/Req/ImageMagick.7.1.1.Q16-HDRI.64-bit.msi"
        filename = "ImageMagickInstaller.msi"

        response = requests.get(url)
        with open(filename, "wb") as f:
            f.write(response.content)

        try:
            subprocess.run(['msiexec', '/i', filename, '/passive'], check=True)
            os.remove(filename)
        except subprocess.CalledProcessError as e:
            if e.returncode == 1602:
                print("A instalação foi cancelada pelo usuário.")
                os.exit()

    def download_imagemagick(self):
        url = "https://github.com/OneDefauter/Menu_/releases/download/Req/ImageMagick.7.1.1.Q16-HDRI.64-bit.msi"
        filename = "ImageMagickInstaller.msi"

        response = requests.get(url)
        with open(filename, "wb") as f:
            f.write(response.content)

        try:
            subprocess.run(['msiexec', '/i', filename, '/passive'], check=True)
            os.remove(filename)
        except subprocess.CalledProcessError as e:
            if e.returncode == 1602:
                print("A instalação foi cancelada pelo usuário.")
                os.exit()

    def show_update_dialog(self):
        if messagebox.askyesno("Nova Atualização", f"Tem uma nova atualização.\nVersão atual: {version}\nVersão mais recente: {latest_version}\nDeseja atualizar agora?"):
            self.install_newversion()
            messagebox.showinfo("Atualização Concluída", "A atualização foi concluída com sucesso!")
            self.root.destroy()

    def check_for_updates(self):
        if latest_version is not None:
            if version != latest_version:
                self.show_update_dialog()
        return True
        
    def install_newversion(self):
        update_url = f"https://github.com/OneDefauter/Juntar-Imagens/archive/refs/tags/{latest_version}.zip"
        response = requests.get(update_url)
        response.raise_for_status()
        if os.path.exists(f"{latest_version}.zip"):
            os.remove(f"{latest_version}.zip")
        with open(f"{latest_version}.zip", "wb") as f:
            f.write(response.content)

        with zipfile.ZipFile(f"{latest_version}.zip", 'r') as zip_ref:
            zip_ref.extractall()

        if os.path.exists(f"Juntar-Imagens-{latest_version.replace('v', '')}/.gitignore"):
            os.remove(f"Juntar-Imagens-{latest_version.replace('v', '')}/.gitignore")
        os.remove(f"Juntar-Imagens-{latest_version.replace('v', '')}/app.py")
        os.remove("app.exe")
        shutil.move(f"Juntar-Imagens-{latest_version.replace('v', '')}/app.exe", self.current_dir)
        os.removedirs(f"Juntar-Imagens-{latest_version.replace('v', '')}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageCropperApp(root)
    app.check_for_updates()
    root.mainloop()