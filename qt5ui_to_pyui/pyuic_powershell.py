import os
os.system("python -m PyQt5.uic.pyuic -x .\\ui\\main_window.ui -o .\\ui\\main_window.py")
os.system("python -m PyQt5.uic.pyuic -x .\\ui\\login_stack.ui -o .\\ui\\login_stack.py")
os.system("python -m PyQt5.uic.pyuic -x .\\ui\\add_inproject_dialog.ui -o .\\ui\\add_inproject_dialog.py")
os.system("python -m PyQt5.uic.pyuic -x .\\ui\\add_new_user_dialog.ui -o .\\ui\\add_new_user_dialog.py")
os.system("python -m PyQt5.uic.pyuic -x .\\ui\\add_new_project_dialog.ui -o .\\ui\\add_new_project_dialog.py")