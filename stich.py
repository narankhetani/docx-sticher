import os
from kivy.app import App
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.lang.builder import Builder
from docxcompose.composer import Composer
from docx import Document as Document_compose
from docx import Document
from kivy.uix.label import Label
from kivy.uix.widget import Widget
from kivy.uix.floatlayout import FloatLayout 

from os.path import expanduser
import os.path

Builder.load_string('''
<StichMessagePopup@Popup>:
    title: 'Sticher Ready'
    size_hint: None, None
    size: 600, 400

    BoxLayout:
        orientation: "vertical"
        Label:
            text: 'Ready to stich selected folder?'
        BoxLayout:
            size_hint_y: 0.5
            Button:
                text: "Cancel"
                on_release: root.dismiss()
            Button:
                text: "Accept"
                on_release:
                    root.parent_inst.stich()
                    root.dismiss()
''')

def combine_word_documents(selectedPath, files):
    """
    :param files: an iterable with full paths to docs
    :return: a Document object with the merged files
    """
    for filnr, fname in enumerate(files):
        file = f"{selectedPath}/{fname.get('filename')}"
        if filnr == 0:
            merged_document = Document(file)
            merged_document.add_page_break()
        else:
            sub_doc = Document(file)

            # Don't add a page break if you've reached the last file.
            if filnr < len(files)-1:
                sub_doc.add_page_break()

            for element in sub_doc.element.body:
                merged_document.element.body.append(element)

    return merged_document

class StichMessagePopup(Popup):
    def __init__(self, parent_inst, *args,  **kwargs):
        super(StichMessagePopup, self).__init__(*args, **kwargs)
        self.parent_inst = parent_inst

class MainWindow(BoxLayout):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.orientation = "vertical"
        self.fichoo = FileChooserListView(path=expanduser('~'), filters=['*.docx'])
        self.popup = StichMessagePopup(self)
        btn_stich = Button(text="Stich", on_release=self.popup.open, size_hint_y=0.1)

        self.add_widget(self.fichoo)
        self.add_widget(btn_stich)

    def show_popup(self, title, message):
        self.box=FloatLayout()

        lab=(
            Label(
                text=message,
                size=(400,300)
        ))
        okayButton = (
            Button(
                text = "Okay",
                size_hint = (0.215, 0.075)
        ))

        self.box.add_widget(lab)
        self.box.add_widget(okayButton)

        self.pop_up = Popup(
            title=title, 
            content=self.box,
            size=(450,300)
        )
        okayButton.bind(on_press=self.pop_up.dismiss)
        self.pop_up.open()


        # self.box=FloatLayout()
        
    #     self.lab=(Label(text="this is a pop exampleby vitriol ",font_size=15, size_hint=(None,None),pos_hint={'x':.25,'y':.6}))
        # self.but=(Button(text="close",size_hint=(None,None),
        # 	width=200,height=50,pos_hint={'x':0,'y':0}))
        # self.box.add_widget(self.but)
        
        # self.box.add_widget(Button(text="blank",size_hint=(None,None),
        # 	width=200,height=50,pos_hint={'x':.5,'y':0}))
       
        # self.main_pop = Popup(title="my_pop",content=self.box,
        # 	size_hint=(None,None),size=(450,300),auto_dismiss=False,title_size=15)
        	
        # self.but.bind(on_press=self.main_pop.dismiss)
        
        # self.main_pop.open()


    def stich(self, *args):
        selectedPath = self.fichoo.path
        filesToBeStiched = []
        content = Button(text='Close me!')

        if os.path.isfile(f"{selectedPath}/merged.docx"):
            self.show_popup('Merge Failed','Merged file already exists, please delete')
            return

        try:
            for root, dirs, files in os.walk(selectedPath):
                for filename in files:
                    if "(" not in filename and "docx" in filename:
                        try:
                            number = filename.split("_").pop().split(".")[0].lstrip('0')
                            filesToBeStiched.append({"number":number, "filename": filename})
                        except Exception as e:
                            print("wasnt able to get number for: "+filename+str(e))

            orderedFiles = sorted(filesToBeStiched, key=lambda k: int(k['number']))
            doc=combine_word_documents(selectedPath, orderedFiles)
            doc.save(f"{selectedPath}/merged.docx")
            self.show_popup('Merge Completed',f"saved to: {selectedPath}/merged.docx")
            return
        except Exception as oe:
            self.show_popup('Merge Failed',f"Reason: {str(oe)}")
            return

        self.fichoo._update_files()


class SticherApp(App):
    def build(self):
        return MainWindow()

if __name__ == "__main__":
    SticherApp().run()