"""GUI program to interact with ETABS to calculate del_ns"""

from tkinter import Button, Tk, HORIZONTAL,Label,Label,Entry,Scale,LabelFrame,messagebox,DISABLED,NORMAL
import sys
import os
import comtypes.client
import math
import pandas as pd
import shutil
import time
from operator import itemgetter

class Input(Tk):
    def __init__(self):
        super().__init__() # initialise the superclass Tk

        self.title("Del_ns")
        self.iconbitmap("icon.ico")
        frame = LabelFrame(self,text="Input",padx = 5,pady = 5)
        frame.grid(row=1,column=1)

        Label(frame,text = "f_ck (MPa) =").grid(row = 0,column = 0)
        Label(frame,text = "\ndel_ns (upper limit) =").grid(row = 1,column = 0)

        self.entry1 = Entry(frame,width=20)
        self.entry1.insert(0,40)
        self.entry1.grid(row = 0,column = 1)
        self.entry2 = Scale(frame,from_ = 1,to = 2,orient = HORIZONTAL,resolution=0.1) # our del_ns only for >1
        self.entry2.set(1.4)
        self.entry2.grid(row = 1,column = 1)

        self.button = Button(frame,text = "OK",width=8,relief = 'raised')
        self.button.bind('<Button-1>', self.assign)
        self.bind('<Return>', self.assign)
        self.button.grid(row = 2,column=0,columnspan = 2,padx=10,pady=10)
        
        self.attach_to_instance()
            
    def attach_to_instance(self):
        try:
            #get the active ETABS object
           self.myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject") 
        except (OSError, comtypes.COMError):
            self.no_model()

    def backup(self,file_path):
        # model backup
        # SapModel.File.Save(file_path)
        os.chdir(os.path.dirname(file_path))
        file_name_ext = os.path.basename(file_path)
        file_name,ext = os.path.splitext(file_name_ext)
        time_stamp = time.strftime("%Y%m%d-%H%M%S")
        new_file_name = file_name+ "_" + time_stamp + ext
        try:
            os.mkdir(".//_backup")
        except FileExistsError:
            pass
        os.chdir(".//_backup")
        shutil.copy2(file_path,new_file_name)

    def del_ns(self,SapModel):
        error_control = 0.03
        fck = float(self.entry1.get())
        thresh = float(self.entry2.get()) - error_control
        es = 200000000 # modulus of elasticity of steel in kN/m2
        col_cover_eff = 100 # i kow its not 100 but this value gives accurate results
        ec = 4700 * math.sqrt(fck) # modulus of elasticity of concrete N/mm2
        ec = ec * 1000 # ec in kn/m2

        #assumptions
        k = 1 # assumed slenderness factor as 1 as this is the worst case for non sway scenario
        beta_dns = 0.6 # code recommended value is 0.6
        col_cover_eff = 0.1

        #===============================================================================================================
        SapModel.SetPresentUnits_2(4,6,2) # kN m C
        SapModel.SetPresentUnits(6) #kn_m_C
        SapModel.SelectObj.ClearSelection() 
        
        #run model (this will create the analysis model)
        SapModel.Analyze.RunAnalysis()

        #===============================================================================================================
        section_data = SapModel.PropFrame.GetAllFrameProperties_2()[1:-1] # transposing data
        section_data = pd.DataFrame.from_records(section_data,).T
        section_data.columns = ["Section","Property Type Enum","t3","t2","tf","tw","t2b","tfb","Area"]
        #===============================================================================================================
        rebar_data  = []
        for section in section_data.Section:
            *_,numb_3,numb_2,bar_size,_,_,_,_,_,_ = SapModel.PropFrame.GetRebarColumn(section)
            rebar_data.append([section,numb_3,numb_2,bar_size])
        rebar_data = pd.DataFrame.from_records(rebar_data,columns = ["Section","Long bars in 3 direction",
                                                                                "Long bars in 2 direction","Bar_dia"])
        section_data = pd.merge(section_data,rebar_data,on = "Section").drop(["Property Type Enum",
                                                                                "tf","tw","t2b","tfb"],axis = 1)
        #===============================================================================================================
        prop_frame_link = []
        for label in SapModel.FrameObj.GetLabelNameList()[1]:
            if SapModel.FrameObj.GetDesignOrientation(label)[0] == 1: # we are only intersted in columns
                prop_frame_link.append([label,SapModel.FrameObj.GetSection(label)[0]])
        prop_frame_link = pd.DataFrame.from_records(prop_frame_link)
        prop_frame_link.columns = ["Unique_Label","Section"]
        frame_data = pd.merge(section_data,prop_frame_link,on = "Section")
        frame_data = frame_data.set_index("Unique_Label")
        frame_data = frame_data.dropna()
        frame_data.Bar_dia = frame_data.Bar_dia.astype('int32') # for some reason bar dia is str
        frame_data.Bar_dia = frame_data.Bar_dia / 1000 # for some reason bar dia is always in mm
        #===============================================================================================================
        bar_area = math.pi / 4 * frame_data.Bar_dia ** 2
        ig_22 = frame_data.t3 * pow(frame_data.t2,3) / 12 # gross moment of inertia in 22 direction
        ig_33 = frame_data.t2 * pow(frame_data.t3,3) / 12 # gross moment of inertia in 33 direction
        def bar_MI(dist, area): return area * pow(dist, 2)
        # moment of inertia of bars
        i_se_22 = 2 * frame_data.loc[:,"Long bars in 2 direction"] * bar_MI(frame_data.t2/2 - col_cover_eff,bar_area) 
        i_se_33 = 2 * frame_data.loc[:,"Long bars in 3 direction"] * bar_MI(frame_data.t3/2 - col_cover_eff,bar_area)
        # etabs uses less accurate version of this equation.
        frame_data["ei_eff_22"] = (0.2 * ec * ig_22 + es * i_se_22)/(1 + beta_dns)
        frame_data["ei_eff_33"] = (0.2 * ec * ig_33 + es * i_se_33)/(1 + beta_dns)
        #===============================================================================================================
        cur_code = SapModel.DesignConcrete.GetCode()[0]
        # it has been found that del_ns is critical for unbraced length > 1.
        # So we only calculate del_ns for those frames
        SapModel.DesignConcrete.SetCode("ACI 318-08") # catching over write for ACI - 11 not defined in python
        problem_frames = []
        for frame in frame_data.index:
            # checking if unbraced length is program determined or user defined
            if SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 3)[1] and \
                                                    SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 4)[1]:
                # catching frames with more than 1 unbraced lengrh
                if (SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 3)[0] > 1) and \
                                            (SapModel.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 4)[0] > 1):
                    problem_frames.append(frame)
        #===============================================================================================================
        f = itemgetter(1,2,5,8)
        ObjectElm = 0
        NumberResults = 0
        data = []
        for frame in problem_frames:
            force_data = SapModel.Results.FrameForce(frame, ObjectElm, NumberResults)
            force_data = pd.DataFrame.from_records(f(force_data)).T
            force_data.columns = ["Unique_Label","Station","Combo","P"]
            temp_data = pd.merge(frame_data,force_data,on = "Unique_Label")
            column_unsupported = temp_data.Station.max()# assuming height is in meter
            pc_22 = (math.pi ** 2 * temp_data.ei_eff_22) / (k * column_unsupported) ** 2
            pc_33 = (math.pi ** 2 * temp_data.ei_eff_33) / (k * column_unsupported) ** 2
            # we are only caluclating del_ns for those frames which is critical(unbraced length > 1). 
            # for such cases Cm is 1
            temp_data["del_ns_22"] = 1 / (1 - temp_data.P.abs()/(0.75 * pc_22))
            temp_data["del_ns_33"] = 1 / (1 - temp_data.P.abs()/(0.75 * pc_33))
      
            thresh_data = temp_data[(temp_data["del_ns_22"] > thresh) | (temp_data["del_ns_33"] > thresh)]
            data.append(thresh_data)
        #===============================================================================================================
        if len(data) == 0:
            messagebox.showinfo(title = "No concrete columns",
                                message = "No concrete columns were found in the active file")
            self.exit()
        thresh_data = pd.concat(data)
        if thresh_data.empty:
            messagebox.showinfo(title = "All columns are safe",
                                message = "All columns have del_ns less than {}".format(thresh + error_control))
            self.exit()
        problem_frames = thresh_data.Unique_Label.unique()
        #===============================================================================================================
        for frame in problem_frames:
            SapModel.FrameObj.SetSelected(frame,True)
        #===============================================================================================================
        # we need to reset our code back to ACI-14
        SapModel.DesignConcrete.SetCode(cur_code)
        SapModel.View.RefreshView(0)

    def assign(self,event):
        """This function is called only when ok button is pressed""" 

        SapModel = self.myETABSObject.SapModel
        file_path = SapModel.GetModelFilename()
        base_name = os.path.basename(file_path)[:-4]

        self.button["state"] = DISABLED
        yes = messagebox.askyesno(title = "Active file is {0}".format(base_name),
                    message = "Is {0} your active model".format(base_name))
        if not yes:
            self.no_model()

        self.backup(file_path) # backup function

        self.lbl = Label(self,text = "Backup created in file root directory")
        self.lbl.grid(row=3,column=0,columnspan=2)
        self.update() # to show above text in window

        self.del_ns(SapModel) # heart of program

        self.lbl.destroy()
        yes = messagebox.askyesno(title = "Failing columns selected",
        message = "Do you wish to continue?")
        if not yes:
            self.exit()
        else:
            self.lbl.destroy()
            self.button["state"] = NORMAL
        

    def no_model(self):
        self.button["state"] = DISABLED
        messagebox.showwarning(title = "Active model not found",
                           message = "Close all ETABS instances and reopen target file")
        self.exit()
        
    def exit(self):
        messagebox.showinfo(title = "Exiting application",message = "For trouble shooting contact me through sbz5677@gmail.com ")
        self.destroy()
        sys.exit()
        
        

if __name__ == '__main__':
    app = Input()
    app.mainloop()

