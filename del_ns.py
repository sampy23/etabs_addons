"""GUI program to interact with ETABS to calculate del_ns"""

from tkinter import Button, Tk, HORIZONTAL,Label,Label,Entry,Scale,LabelFrame,messagebox,DISABLED,NORMAL
from sys import exit
import os
import comtypes.client
from math import pi
import pandas as pd
from shutil import copy2
from time import strftime
from operator import itemgetter

class Input(Tk):
    def __init__(self):
        super().__init__() # initialise the superclass Tk

        self.attach_to_instance()

        self.title("Del_ns")
        self.iconbitmap("icon.ico")
        self.font_size = ("Courier", 16)
        self.row = 0
        self.thresh_input() 

    def attach_to_instance(self):
        try:
            #get the active ETABS object
            self.myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject") 
        except (OSError, comtypes.COMError):
            self.no_model()

    def no_model(self):
        self.withdraw()
        messagebox.showwarning(title = "Active model not found",
                           message = "Close all ETABS instances if any open and reopen the target file first.")
        self.exit() 

    def thresh_input(self):
        windo_size = "600x220"
        height = int(windo_size[4:])
        width = int(windo_size[:3])
        self.geometry(windo_size)
        
        self.frame_1 = LabelFrame(self,height=height,width=width)
        self.frame_2 = LabelFrame(self,height=height,width=width,text="Input",padx = 5,pady = 5)
        self.frame_1.grid(row=0,column=0)
        self.frame_1.grid_propagate(0) # this is for a fixed frame size
        self.frame_2.place(in_=self.frame_1,anchor="c", relx=.5, rely=.5)

        self.label = Label(self.frame_2,text = "\ndel_ns (upper limit) =")
        self.label.grid(row = 0,column = 0)
        self.label.config(font=self.font_size)

        # our del_ns calc only for > 1
        self.entry2 = Scale(self.frame_2,from_ = 1,to = 2,orient = HORIZONTAL,resolution=0.1) 
        self.entry2.set(1.4)
        self.entry2.grid(row = 0,column = 1)

        self.button = Button(self.frame_2,text = "OK",width=8,relief = 'raised')
        self.button.bind('<Button-1>', self.assign)
        self.bind('<Return>', self.assign)
        self.button.grid(row = 1,column=0,columnspan = 2,padx=10,pady=10)
        self.button.config(font=self.font_size)

    def label_fn(self,text):
        self.lbl = Label(self.frame_1,text = text,width = 50,anchor="w",)
        self.lbl.grid(row = self.row,column=0)
        self.lbl.config(font=self.font_size)
        self.update() # to show above text in window
        self.row += 1
        return self.lbl

    def assign(self,event):
        """This function is called only when ok button is pressed""" 
        try:
            self.SapModel = self.myETABSObject.SapModel
        except (OSError, comtypes.COMError):
            self.no_model()

        file_path = self.SapModel.GetModelFilename()
        base_name = os.path.basename(file_path)[:-4]
        self.thresh = float(self.entry2.get()) 
        self.frame_2.destroy()
        self.lbl_1 = self.label_fn("Active file is {0}.".format(base_name))
        self.backup(file_path) # backup function
        self.lbl_2 = self.label_fn("Backup created in file root directory.")
        self.del_ns() # heart of program

    def backup(self,file_path):
        # model backup
        # self.SapModel.File.Save(file_path)
        os.chdir(os.path.dirname(file_path))
        file_name_ext = os.path.basename(file_path)
        file_name,ext = os.path.splitext(file_name_ext)
        time_stamp = strftime("%Y%m%d-%H%M%S")
        new_file_name = file_name+ "_" + time_stamp + ext
        try:
            os.mkdir(".//_backup")
        except FileExistsError:
            pass
        os.chdir(".//_backup")
        copy2(file_path,new_file_name)

    def del_ns(self):
        #assumptions
        beta_dns =  1# code recommended value is 0.6
        #===============================================================================================================
        self.curr_unit = self.SapModel.GetPresentUnits()
        self.SapModel.SetPresentUnits(6) #kn_m_C
        self.SapModel.SelectObj.ClearSelection() 
        #===============================================================================================================
        #run model (this will create the analysis model)
        self.lbl_analysis = self.label_fn("Analysing ........................")
        self.SapModel.Analyze.RunAnalysis()
        self.lbl_analysiscomplete = self.label_fn("Analyses complete.")
        #===============================================================================================================
        # selecting load cases for output. Otherwise error will be generated for self.SapModel.Results.FrameForce
        _,combos,_ = self.SapModel.RespCombo.GetNameList(1, " ")
        self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        combos = [x for x in combos if x.startswith("U") and not x.endswith("O")]
        for combo in combos:
            self.SapModel.Results.Setup.SetComboSelectedForOutput(combo,True) 
        #===============================================================================================================
        section_data = self.SapModel.PropFrame.GetAllFrameProperties_2()[1:-1] # transposing data
        section_data = pd.DataFrame.from_records(section_data,).T
        section_data.columns = ["Section","Property Type Enum","t3","t2","tf","tw","t2b","tfb","Area"]
        #===============================================================================================================
        prop_frame_link = []
        for label in self.SapModel.FrameObj.GetLabelNameList()[1]:
            if self.SapModel.FrameObj.GetDesignOrientation(label)[0] == 1: # we are only intersted in columns
                prop_frame_link.append([label,self.SapModel.FrameObj.GetSection(label)[0]])

        if len(prop_frame_link) == 0:
            self.lbl_3 = self.label_fn("No columns were found in the active file.")
            self.exit()
        else:
            self.lbl_3 = self.label_fn("{0} columns found in the model.".format(len(prop_frame_link)))

        prop_frame_link = pd.DataFrame.from_records(prop_frame_link)
        prop_frame_link.columns = ["Unique_Label","Section"]
        frame_data = pd.merge(section_data,prop_frame_link,on = "Section")
        frame_data = frame_data.set_index("Unique_Label")
        #===============================================================================================================
        ig_22 = frame_data.t3 * pow(frame_data.t2,3) / 12 # gross moment of inertia in 22 direction
        ig_33 = frame_data.t2 * pow(frame_data.t3,3) / 12 # gross moment of inertia in 33 direction

        def section_fck(df,df_column):
            ls = []
            for section in df_column:
                fck_string = self.SapModel.PropFrame.GetMaterial(section)[0]
                fck = self.SapModel.PropMaterial.GetOConcrete(fck_string)[0]/1000 # we want in MPa
                ls.append(fck)
            df["fck"] = ls
            return df

        frame_data = section_fck(frame_data,frame_data["Section"])
        ec = 4700 *frame_data["fck"].pow(1/2) * 1000 # ec in kn/m2
        # etabs preferred equation.
        frame_data["ei_eff_22"] = (0.4 * ec * ig_22)/(1 + beta_dns)
        frame_data["ei_eff_33"] = (0.4 * ec * ig_33)/(1 + beta_dns)
        #===============================================================================================================
        cur_code = self.SapModel.DesignConcrete.GetCode()[0]
        self.SapModel.DesignConcrete.SetCode("ACI 318-08") # catching over write for ACI - 11 not defined in python
        problem_frames = []
        # the idea is that column with buckling issue never have del_ns < 1
        for frame in frame_data.index:
            # catching frames with more than 1 unbraced length
            # this will also filter out all steel columns
            if (self.SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 3)[0] >= 1) or \
                                            (self.SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 4)[0] >= 1):
                problem_frames.append(frame)
        #===============================================================================================================
        f = itemgetter(1,2,5,8)
        ObjectElm = 0
        NumberResults = 0
        data = []
        for frame in problem_frames:
            force_data = self.SapModel.Results.FrameForce(frame, ObjectElm, NumberResults)
            force_data = pd.DataFrame.from_records(f(force_data)).T
            force_data.columns = ["Unique_Label","Station","Combo","P"]
            temp_data = pd.merge(frame_data,force_data,on = "Unique_Label")
            # end length offset has to be added if present
            # assuming height is in meter
            column_length = temp_data.Station.max() + self.SapModel.FrameObj.GetEndLengthOffset(frame)[2]
            unbrac_minor = self.SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 4)[0]
            unbrac_major = self.SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 3)[0]
            column_unsupported_minor = unbrac_minor * column_length
            column_unsupported_major = unbrac_minor * column_length
            pc_22 = (pi ** 2 * temp_data.ei_eff_22) / (1 * column_unsupported_minor) ** 2
            pc_33 = (pi ** 2 * temp_data.ei_eff_33) / (1 * column_unsupported_major) ** 2
            # Calculation of Cm is little obscure for etabs data as it tends to get muddled
            # so for a conservative approach we take Cm as 1
            temp_data["del_ns_22"] = 1 / (1 - temp_data.P.abs()/(0.75 * pc_22))
            temp_data["del_ns_33"] = 1 / (1 - temp_data.P.abs()/(0.75 * pc_33))
      
            thresh_data = temp_data[(temp_data["del_ns_22"] > self.thresh) | (temp_data["del_ns_33"] > self.thresh)]
            data.append(thresh_data)
        #===============================================================================================================
        # no concrete columns
        if len(data) == 0:
            self.lbl_4 = self.label_fn("No concrete columns found in the active model.")
            self.cont_yesno()
        # concrete columns found
        else:
            thresh_data = pd.concat(data)
            if thresh_data.empty:
                self.lbl_5 = self.label_fn("All columns have del_ns less than {0}".format(self.thresh))
                self.safe = True
                self.cont_yesno()
            else:
                self.safe = False
                problem_frames = thresh_data.Unique_Label.unique()
                #===============================================================================================================
                self.lbl_5 = self.label_fn("{0} columns likely to have buckling issues.".format(len(problem_frames)))
                for frame in problem_frames:
                    self.SapModel.FrameObj.SetSelected(frame,True)
                self.lbl_6 = self.label_fn("Check columns selected in the model.")
                #===============================================================================================================
                # we need to reset our code back to ACI-14
                self.SapModel.DesignConcrete.SetCode(cur_code)
                self.SapModel.View.RefreshView(0)
            if not self.safe:
                self.cont_yesno()

    def cont_yesno(self):
        yes = messagebox.askyesno(title = "Failing columns selected",
        message = "Do you wish to continue?")
        self.lbl_analysis.destroy()
        self.lbl_analysiscomplete.destroy()
        self.lbl_1.destroy()
        self.lbl_2.destroy()
        self.lbl_3.destroy()
        # exception for if concrete columns present
        try:
            self.lbl_4.destroy()
        except:
            pass
        # exception to deal with no concrete columns
        try:
            self.lbl_5.destroy()
            self.lbl_6.destroy()
        except:
            pass

        if not yes:
            self.exit()
        else:
            self.thresh_input() 

    def exit(self):
        # exception for call from "no model"
        text = "Note!!\nThis program calculates Del_ns only for load\ncombinations starting with \"U\" and ending" \
                                                                                    " not\n with \"O\""
        self.lbl = Label(self.frame_1,text = text,width = 50,anchor="w",)
        self.lbl.grid(row = 0,column=0)
        self.lbl.config(font=self.font_size)
        self.update()

        messagebox.showinfo(title = "Help",message = "For trouble shooting contact me through sbz5677@gmail.com ")
        self.destroy()
        self.SapModel.SetPresentUnits(self.curr_unit) 
        exit()
        
if __name__ == '__main__':
    app = Input()
    app.mainloop()