"""GUI program to interact with ETABS to calculate del_ns"""

from tkinter import Button, Tk, HORIZONTAL,Label,Label,Entry,Scale,LabelFrame,messagebox,DISABLED,NORMAL,IntVar
from sys import exit
import os
import comtypes.client
from math import pi
import pandas as pd
from shutil import copy2
import time
from operator import itemgetter

class Input(Tk):
    def __init__(self):
        """Initialise the class Input"""
        super().__init__() # initialise the superclass Tk

        self.attach_to_instance()
        self.beta_dns =  1# code recommended value is 0.6

        self.title("Del_ns")
        self.font_size = ("Courier", 12)
        self.thresh_input() 
    def attach_to_instance(self):
        """Attaches to an existing ETABS application"""
        try:
            #get the active ETABS object
            self.myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject") 
        except (OSError, comtypes.COMError):
            self.no_model()
    def no_model(self):
        """To handle no ETABS model open exception"""
        self.withdraw()
        messagebox.showwarning(title = "Active model not found",
                           message = "Close all ETABS instances if any open and reopen the target file first.")
        self.exit() 
    def label_fn_frame_1(self,text,frame = None):
        """Function to push text onto frame 1"""
        self.lbl = Label(self.frame_1,text = text,width = 50,anchor="w")
        self.lbl.grid(row = self.row_1,column=0)
        self.lbl.config(font=self.font_size)
        self.update() # to show above text in window
        self.row_1 += 1
        return self.lbl
    def label_fn_frame_2(self,text,frame = None):
        """Function to push text onto frame 2"""
        self.lbl = Label(self.frame_2,text = text,width = 50)
        self.lbl.grid(row = self.row_2,column=0)
        self.lbl.config(font=self.font_size)
        self.update() # to show above text in window
        self.row_2 += 1
        return self.lbl
    def thresh_input(self):
        """Produces a slider for user to input threshold value for del_ns, frames below which values will be ignored"""
        self.row_1 = 0
        self.row_2 = 1

        self.frame_1 = LabelFrame(self,text="Output")
        self.frame_2 = LabelFrame(self,text="Input")
        self.frame_2.grid(row=0,column=0)

        self.label = Label(self.frame_2,text = "\ndel_ns (upper limit) =")
        self.label.grid(row = 0,column = 0)
        self.label.config(font=self.font_size)

        # our del_ns calc only for > 1
        self.entry1 = Scale(self.frame_2,from_ = 1,to = 2,orient = HORIZONTAL,resolution=0.1) 
        self.entry1.set(1.4)
        self.entry1.grid(row = 0,column = 1)
        # button for slow input
        self.button_fast = Button(self.frame_2,text = "FAST",width=8,relief = 'raised',fg = "green")
        self.button_fast.bind('<Button-1>', self.assign_fast)
        self.bind('<Return>', self.assign_fast)
        self.button_fast.grid(row = 3,column=0,columnspan = 1,padx=10,pady=10)
        self.button_fast.config(font=self.font_size)
        # button for slow input
        self.button_slow = Button(self.frame_2,text = "SLOW",width=8,relief = 'raised',fg = "red")
        self.button_slow.bind('<Button-1>', self.assign_slow)
        self.bind('<Return>', self.assign_slow)
        self.button_slow.grid(row = 3,column=1,columnspan = 1,padx=10,pady=10)
        self.button_slow.config(font=self.font_size)
    def assign_fast(self,event):
        """This function is called only when fast button is pressed""" 

        try:
            self.SapModel = self.myETABSObject.SapModel
        except (OSError, comtypes.COMError):
            self.no_model()

        self.model_path = self.SapModel.GetModelFilename()
        base_name = os.path.basename(self.model_path)[:-4]
        self.thresh = float(self.entry1.get()) 
        self.frame_2.destroy()
        self.frame_1.grid(row=0,column=0)
        self.lbl_1 = self.label_fn_frame_1("Active file is {0}.".format(base_name))
        self.backup(self.model_path) # backup function
        self.lbl_2 = self.label_fn_frame_1("Backup created in file root directory.")
        self.del_ns_fast() # heart of program       
    def assign_slow(self,event):
        """This function is called only when slow button is pressed""" 
        def combo_selected():
            self.load_starts = self.entry2.get().lower()
            self.load_notends = self.entry3.get().lower()
            try:
                self.SapModel = self.myETABSObject.SapModel
            except (OSError, comtypes.COMError):
                self.no_model()

            self.model_path = self.SapModel.GetModelFilename()
            base_name = os.path.basename(self.model_path)[:-4]
            self.thresh = float(self.entry1.get()) 
            self.frame_2.destroy()
            self.frame_1.grid(row=0,column=0)
            self.lbl_1 = self.label_fn_frame_1("Active file is {0}.".format(base_name))
            self.backup(self.model_path) # backup function
            self.lbl_2 = self.label_fn_frame_1("Backup created in file root directory.")
            self.del_ns_slow() # heart of program

        self.button_fast.destroy()
        self.button_slow.destroy()

        # enter the the first few letters of load combination to calculate del_ns.
        self.lbl = self.label_fn_frame_2("Calculate for load combo starting with:")
        self.entry2 = Entry(self.frame_2,width=20)
        self.entry2.insert(1,"U")
        self.entry2.grid(row = self.row_2-1,column=1,columnspan = 1,padx=10,pady=10)
        self.entry2.config(font=self.font_size)
        # enter the the last few letters of load combination to be excluded for o for overstrength combinations
        self.lbl = self.label_fn_frame_2("Exclude load combo ending with:")
        self.entry3 = Entry(self.frame_2,width=20)
        self.entry3.insert(1,"O")
        self.entry3.grid(row = self.row_2-1,column=1,columnspan = 1,padx=10,pady=10)
        self.entry3.config(font=self.font_size)
        # ok button
        self.button = Button(self.frame_2,text = "OK",width=8,relief = 'raised',command=combo_selected)
        self.button.grid(row = 3,column=0,columnspan = 2,padx=10,pady=10)
        self.button.config(font=self.font_size)
 
    def backup(self,model_path):
        """Function to create backup of attached ETABS"""
        if 4 in set(self.SapModel.Analyze.GetCaseStatus()[2]): # 4 indicates presence of run case
            #atleast one case has run so no need to save the file and lose analysis data
            pass
        else:
            self.SapModel.File.Save(model_path) #save file before taking backup, but this deletes program results
        os.chdir(os.path.dirname(model_path))
        # model backup
        try:
            os.mkdir(".//_backup")
        except FileExistsError:
            pass
        file_dir = os.path.dirname(model_path)
        file_name_ext = os.path.basename(model_path)
        file_name,ext = os.path.splitext(file_name_ext)
        time_stamp = time.strftime("%H%M%S")
        new_file_name = file_name+ "_" + time_stamp + ext
        new_model_path = os.path.join("_backup",new_file_name)
        self.new_model_path = os.path.join(file_dir,new_model_path)
        copy2(model_path,new_model_path)  

    def section_fck(self,df,df_column):
        ls = []
        for section in df_column:
            fck_string = self.SapModel.PropFrame.GetMaterial(section)[0]
            fck = self.SapModel.PropMaterial.GetOConcrete(fck_string)[0]/1000 # we want in MPa
            ls.append(fck)
        df["fck"] = ls
        return df

    @staticmethod
    def env_cm(end1,end2):
        """When we have EQ cases or envelope cases we will have maximum and minimum cases. 
        In that case we need to combine them, which is not the right thing. Instead we should be working with table
        Column Design Forces to calculate Cm. Unfortunately that table is not accessible through API.
        So we simply make use of what we have and follow the code requirement:
            1. Find absolute maximum at one end and absolute minimum at other end to calculate Cm
        """
        temp = end1 + end2
        abs_max_1 = sorted(temp,reverse = True,key=abs)[0]

        indx = temp.index(abs_max_1)

        if indx < 2:
            search_end = end2
        else:
            search_end = end1

        if abs_max_1 == 0:
            return 1
        
        cm_1 = 0.6 + 0.4 * search_end[0] / abs_max_1
        cm_2 = 0.6 + 0.4 * search_end[1]  / abs_max_1

        cm = max (cm_1,cm_2) 
        return cm

    def apply_cm(self,x):
        """Function that applies a series actions on pandas dataframe to generate Cm values"""
        if len(x) == 6: # for sway cases
            end1 = [x.iloc[0,1],x.iloc[3,1]]
            end2 = [x.iloc[2,1],x.iloc[5,1]]
            cm = self.env_cm(end1,end2)
        else: 
            # this function returns absolute max value in the group by preserving the sign
            sign_abs_max_moment =   max(x.iloc[:,1].min(), x.iloc[:,1].max(), key=abs) 
            sign_abs_min_moment =  min(x.iloc[:,1].min(), x.iloc[:,1].max(), key=abs)
            abs_max_end_moment = x.iloc[[0,2],1].abs().max()
            abs_max_middle_moment = abs(x.iloc[1,1])
            if abs_max_end_moment <= abs_max_middle_moment:
                cm = 1 # if end moments are lesser no correction need to applied
            else:
                cm = 0.6 + 0.4 * sign_abs_min_moment/sign_abs_max_moment
                
        x.loc[:,"CM"] = cm
        return x
    def del_ns_fast(self):
        """Calculation of del_ns only for those combination which produce maximum PMM ratio in frame, hence much faster"""
        #===============================================================================================================
        self.curr_unit = self.SapModel.GetPresentUnits()
        self.SapModel.SetPresentUnits(6) #kn_m_C
        self.SapModel.SelectObj.ClearSelection() 
        #===============================================================================================================
        #run model (this will create the analysis model)
        self.lbl_analysis = self.label_fn_frame_1("Analysing ........................")
        self.SapModel.Analyze.RunAnalysis()
        if not self.SapModel.DesignConcrete.GetResultsAvailable():
            self.SapModel.DesignConcrete.StartDesign()
        self.lbl_analysiscomplete = self.label_fn_frame_1("Analyses complete.")
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
            self.lbl_3 = self.label_fn_frame_1("No columns were found in the active file.")
            self.exit()
        else:
            self.lbl_3 = self.label_fn_frame_1("{0} columns found in the model.".format(len(prop_frame_link)))
            self.lbl_3 = self.label_fn_frame_1("Calculating del_ns............")
            #start time
            start = time.time()

        prop_frame_link = pd.DataFrame.from_records(prop_frame_link)
        prop_frame_link.columns = ["Unique_Label","Section"]
        frame_data = pd.merge(section_data,prop_frame_link,on = "Section")
        frame_data = frame_data.set_index("Unique_Label")
        #===============================================================================================================
        ig_22 = frame_data.t3 * pow(frame_data.t2,3) / 12 # gross moment of inertia in 22 direction
        ig_33 = frame_data.t2 * pow(frame_data.t3,3) / 12 # gross moment of inertia in 33 direction

        frame_data = self.section_fck(frame_data,frame_data["Section"])
        ec = 4700 *frame_data["fck"].pow(1/2) * 1000 # ec in kn/m2
        # etabs preferred equation.
        frame_data["ei_eff_22"] = (0.4 * ec * ig_22)/(1 + self.beta_dns)
        frame_data["ei_eff_33"] = (0.4 * ec * ig_33)/(1 + self.beta_dns)
        frame_data = frame_data[frame_data.ei_eff_22 != 0] # filtering out steel columns
        #===============================================================================================================
        # collecting result data before saving as model. This action is required because we are setting a different code.
        # setting and resetting code affects design preferences. Workaround this problem is collect all data
        # set code and then save as new model to collect unbraced length and then return to original model
        f = itemgetter(1,2,5,8,12,13)
        f_pmm = itemgetter(1,4,6)
        ObjectElm = 0
        NumberResults = 0
        force_data_list = []
        for frame in frame_data.index:
            self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            col_data = self.SapModel.DesignConcrete.GetSummaryResultsColumn(frame)
            col_data = pd.DataFrame.from_records(f_pmm(col_data)).T
            col_data.columns = ["Unique_Label","Combo","PMM"]
            combo = list(col_data[col_data.PMM == col_data.PMM.max()].Combo)[0]
            self.SapModel.Results.Setup.SetComboSelectedForOutput(combo,True) 

            force_data = self.SapModel.Results.FrameForce(frame, ObjectElm, NumberResults)
            force_data = pd.DataFrame.from_records(f(force_data)).T
            force_data.columns = ["Unique_Label","Station","Combo","P","M2","M3"]
            force_data_list.append(force_data)
        force_data = pd.concat(force_data_list,axis=0)    
        #===============================================================================================================
        ProgramPath = r"C:\Program Files\Computers and Structures\ETABS 17\ETABS.exe"
        #create API helper object
        helper = comtypes.client.CreateObject('ETABSv17.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv17.cHelper)
        self.myETABSObject_dummy = helper.CreateObject(ProgramPath)
        #start ETABS application
        self.myETABSObject_dummy.ApplicationStart()
        #Hide the current program
        self.myETABSObject_dummy.Hide() 
        
        #create SapModel object
        self.SapModel_dummy = self.myETABSObject_dummy.SapModel
        self.SapModel_dummy.File.OpenFile(self.new_model_path) #get back to our model
        self.SapModel_dummy.SetPresentUnits(6) #kn_m_C
        self.SapModel_dummy.DesignConcrete.SetCode("ACI 318-08") # catching over write for ACI - 11 not defined in python
        #===============================================================================================================
        frame_force_data = pd.merge(frame_data,force_data,on = "Unique_Label")
        data = []
        for frame in frame_data.index:
            temp_data = frame_force_data[frame_force_data.Unique_Label == frame].copy() #to avoid SettingWithCopyWarning
            # end length offset has to be added if present
            # assuming height is in meter
            column_length = temp_data.Station.max() + self.SapModel_dummy.FrameObj.GetEndLengthOffset(frame)[2]
            unbrac_minor = self.SapModel_dummy.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 4)[0]
            unbrac_major = self.SapModel_dummy.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 3)[0]
            column_unsupported_minor = unbrac_minor * column_length
            column_unsupported_major = unbrac_major * column_length
            pc_22 = (pi ** 2 * temp_data.ei_eff_22) / (1 * column_unsupported_minor) ** 2
            pc_33 = (pi ** 2 * temp_data.ei_eff_33) / (1 * column_unsupported_major) ** 2
            temp_data.loc[:,"p_critical22"] = pc_22
            temp_data.loc[:,"p_critical33"] = pc_33
            temp_data.loc[:,"CM22"] = temp_data.groupby("Combo")[["Station","M2"]].apply(self.apply_cm).CM
            temp_data.loc[:,"CM33"] = temp_data.groupby("Combo")[["Station","M3"]].apply(self.apply_cm).CM
            # so for a conservative approach we take Cm as 1
            temp_data.loc[:,"del_ns_22"] = temp_data["CM22"] / (1 - temp_data.P.abs()/(0.75 * pc_22))
            temp_data.loc[:,"del_ns_33"] = temp_data["CM33"] / (1 - temp_data.P.abs()/(0.75 * pc_33))
            # minimum value of del_ns is 1
            temp_data.loc[temp_data.del_ns_22 < 1,"del_ns_22"] = 1
            temp_data.loc[temp_data.del_ns_33 < 1,"del_ns_33"] = 1
            thresh_data = temp_data[(temp_data["del_ns_22"] > self.thresh) | (temp_data["del_ns_33"] > self.thresh)]
            data.append(thresh_data)
        #===============================================================================================================
        # get rid of hodden file as fast as you can
        self.myETABSObject_dummy.ApplicationExit(False) # we donot wish to save the backup file
        # no concrete columns
        if len(data) == 0:
            self.lbl_4 = self.label_fn_frame_1("No concrete columns found in the active model.")
            self.cont_yesno()
        # concrete columns found
        else:
            thresh_data = pd.concat(data)
            thresh_data = thresh_data.drop(["Property Type Enum","t3","t2","tf","tw","t2b","tfb","Area","fck",\
                                                                                    "ei_eff_22","ei_eff_33"],axis = 1)
            # for visual checking
            try:
                with pd.ExcelWriter("DEL_NS.xlsx") as writer:
                    thresh_data.to_excel(writer,index = False)
            except PermissionError:
                text = "\"DEL_NS.xlsx\" cannot be generated as it is open"
                self.file_open = self.label_fn_frame_1(text)
                self.file_open.config(fg="red")

            if thresh_data.empty:
                self.lbl_5 = self.label_fn_frame_1("All columns have del_ns less than {0}".format(self.thresh))
                self.lbl_5.config(fg="green")
                self.safe = True
                self.cont_yesno()
            else:
                self.safe = False
                problem_frames = thresh_data.Unique_Label.unique()
                #=======================================================================================================
                self.lbl_5 = self.label_fn_frame_1("{0} columns likely to have buckling issues."\
                                                                                        .format(len(problem_frames)))
                self.lbl_5.config(fg="red")
                for frame in problem_frames:
                    self.SapModel.FrameObj.SetSelected(frame,True)
                end = time.time() # end time of core calculation
                self.lbl_6 = self.label_fn_frame_1("Check columns selected in the model.")
                self.lbl_6 = self.label_fn_frame_1("Time taken for core calculation in seconds is {0}"\
                                                                                            .format(round(end-start)))
                #=======================================================================================================
                self.SapModel.View.RefreshView(0)
            if not self.safe:
                self.cont_yesno()        
    def del_ns_slow(self):
        self.curr_unit = self.SapModel.GetPresentUnits()
        self.SapModel.SetPresentUnits(6) #kn_m_C
        self.SapModel.SelectObj.ClearSelection() 
        #===============================================================================================================
        #run model (this will create the analysis model)
        self.lbl_analysis = self.label_fn_frame_1("Analysing ........................")
        self.SapModel.Analyze.RunAnalysis()
        self.lbl_analysiscomplete = self.label_fn_frame_1("Analyses complete.")
        #===============================================================================================================
        # selecting load cases for output. Otherwise error will be generated for self.SapModel.Results.FrameForce
        _,combos,_ = self.SapModel.RespCombo.GetNameList(1, " ")
        self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        combos = \
            [x for x in combos if x.lower().startswith(self.load_starts) and not x.lower().endswith(self.load_notends)]
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
            self.lbl_3 = self.label_fn_frame_1("No columns were found in the active file.")
            self.exit()
        else:
            self.lbl_3 = self.label_fn_frame_1("{0} columns found in the model.".format(len(prop_frame_link)))
            self.lbl_3 = self.label_fn_frame_1("Calculating del_ns............")
            #start time
            start = time.time()

        prop_frame_link = pd.DataFrame.from_records(prop_frame_link)
        prop_frame_link.columns = ["Unique_Label","Section"]
        frame_data = pd.merge(section_data,prop_frame_link,on = "Section")
        frame_data = frame_data.set_index("Unique_Label")
        #===============================================================================================================
        ig_22 = frame_data.t3 * pow(frame_data.t2,3) / 12 # gross moment of inertia in 22 direction
        ig_33 = frame_data.t2 * pow(frame_data.t3,3) / 12 # gross moment of inertia in 33 direction

        frame_data = self.section_fck(frame_data,frame_data["Section"])
        ec = 4700 *frame_data["fck"].pow(1/2) * 1000 # ec in kn/m2
        # etabs preferred equation.
        frame_data["ei_eff_22"] = (0.4 * ec * ig_22)/(1 + self.beta_dns)
        frame_data["ei_eff_33"] = (0.4 * ec * ig_33)/(1 + self.beta_dns)
        frame_data = frame_data[frame_data.ei_eff_22 != 0] # filtering out steel columns

        f = itemgetter(1,2,5,8,12,13)
        ObjectElm = 0
        NumberResults = 0
        force_data_list = []
        #===============================================================================================================
        # collecting result data before saving as model. This action is required because we are setting a different code.
        # setting and resetting code affects design preferences. Workaround this problem is collect all data
        # set code and then save as new model to collect unbraced length and then return to original model
        for frame in frame_data.index:
            force_data = self.SapModel.Results.FrameForce(frame, ObjectElm, NumberResults)
            force_data = pd.DataFrame.from_records(f(force_data)).T
            force_data.columns = ["Unique_Label","Station","Combo","P","M2","M3"]
            force_data_list.append(force_data)
        force_data = pd.concat(force_data_list,axis=0)    
        #===============================================================================================================
        ProgramPath = r"C:\Program Files\Computers and Structures\ETABS 17\ETABS.exe"
        #create API helper object
        helper = comtypes.client.CreateObject('ETABSv17.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv17.cHelper)
        self.myETABSObject_dummy = helper.CreateObject(ProgramPath)
        #start ETABS application
        self.myETABSObject_dummy.ApplicationStart()
        #Hide the current program
        self.myETABSObject_dummy.Hide() 
        
        #create SapModel object
        self.SapModel_dummy = self.myETABSObject_dummy.SapModel
        self.SapModel_dummy.File.OpenFile(self.new_model_path) #get back to our model
        self.SapModel_dummy.SetPresentUnits(6) #kn_m_C
        self.SapModel_dummy.DesignConcrete.SetCode("ACI 318-08") # catching over write for ACI - 11 not defined in python
        #===============================================================================================================
        frame_force_data = pd.merge(frame_data,force_data,on = "Unique_Label")
        data = []
        for frame in frame_data.index:
            temp_data = frame_force_data[frame_force_data.Unique_Label == frame].copy() #to avoid SettingWithCopyWarning
            # end length offset has to be added if present
            # assuming height is in meter
            column_length = temp_data.Station.max() + self.SapModel_dummy.FrameObj.GetEndLengthOffset(frame)[2]
            unbrac_minor = self.SapModel_dummy.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 4)[0]
            unbrac_major = self.SapModel_dummy.DesignConcrete.ACI318_08_IBC2009.GetOverwrite(frame, 3)[0]
            column_unsupported_minor = unbrac_minor * column_length
            column_unsupported_major = unbrac_major * column_length
            pc_22 = (pi ** 2 * temp_data.ei_eff_22) / (1 * column_unsupported_minor) ** 2
            pc_33 = (pi ** 2 * temp_data.ei_eff_33) / (1 * column_unsupported_major) ** 2
            temp_data.loc[:,"p_critical22"] = pc_22
            temp_data.loc[:,"p_critical33"] = pc_33
            temp_data.loc[:,"CM22"] = temp_data.groupby("Combo")[["Station","M2"]].apply(self.apply_cm).CM
            temp_data.loc[:,"CM33"] = temp_data.groupby("Combo")[["Station","M3"]].apply(self.apply_cm).CM
            # so for a conservative approach we take Cm as 1
            temp_data.loc[:,"del_ns_22"] = temp_data["CM22"] / (1 - temp_data.P.abs()/(0.75 * pc_22))
            temp_data.loc[:,"del_ns_33"] = temp_data["CM33"] / (1 - temp_data.P.abs()/(0.75 * pc_33))
            # minimum value of del_ns is 1
            temp_data.loc[temp_data.del_ns_22 < 1,"del_ns_22"] = 1
            temp_data.loc[temp_data.del_ns_33 < 1,"del_ns_33"] = 1
            thresh_data = temp_data[(temp_data["del_ns_22"] > self.thresh) | (temp_data["del_ns_33"] > self.thresh)]
            data.append(thresh_data)
        #===============================================================================================================
        # get rid of hodden file as fast as you can
        self.myETABSObject_dummy.ApplicationExit(False) # we donot wish to save the backup file
        # no concrete columns
        if len(data) == 0:
            self.lbl_4 = self.label_fn_frame_1("No concrete columns found in the active model.")
            self.cont_yesno()
        # concrete columns found
        else:
            thresh_data = pd.concat(data)
            thresh_data = thresh_data.drop(["Property Type Enum","t3","t2","tf","tw","t2b","tfb","Area","fck",\
                                                                                    "ei_eff_22","ei_eff_33"],axis = 1)
            # for visual checking
            try:
                with pd.ExcelWriter("DEL_NS.xlsx") as writer:
                    thresh_data.to_excel(writer,index = False)
            except PermissionError:
                text = "\"DEL_NS.xlsx\" cannot be generated as it is open"
                self.file_open = self.label_fn_frame_1(text)
                self.file_open.config(fg="red")

            if thresh_data.empty:
                self.lbl_5 = self.label_fn_frame_1("All columns have del_ns less than {0}".format(self.thresh))
                self.lbl_5.config(fg="green")
                self.safe = True
                self.cont_yesno()
            else:
                self.safe = False
                problem_frames = thresh_data.Unique_Label.unique()
                #=======================================================================================================
                self.lbl_5 = self.label_fn_frame_1("{0} columns likely to have buckling issues."\
                                                                                        .format(len(problem_frames)))
                self.lbl_5.config(fg="red")
                for frame in problem_frames:
                    self.SapModel.FrameObj.SetSelected(frame,True)
                end = time.time() # end time of core calculation
                self.lbl_6 = self.label_fn_frame_1("Check columns selected in the model.")
                self.lbl_6 = self.label_fn_frame_1("Time taken for core calculation in seconds is {0}"\
                                                                                            .format(round(end-start)))
                #=======================================================================================================
                self.SapModel.View.RefreshView(0)
            if not self.safe:
                self.cont_yesno()
    def cont_yesno(self):
        yes = messagebox.askyesno(title = "Failing columns selected",
        message = "Do you wish to continue?")
        self.lbl_analysis.destroy()
        self.lbl_analysiscomplete.destroy()
        # exception if files is closed
        try:
            self.file_open.destroy()
        except:
            pass
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
            self.frame_1.destroy()
            self.thresh_input() 
    def exit(self):
        # exception for call from "no model"
        self.destroy()
        self.SapModel.SetPresentUnits(self.curr_unit) 
        exit()
        
if __name__ == '__main__':
    app = Input()
    app.mainloop()