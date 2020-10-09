from collections import defaultdict
import sys
import tkinter as tk
from tkinter import messagebox

class App():
    def __init__(self,master):
        self.master = master
        self.font_size = ("Courier", 12)
        self.master.frame_1 = tk.LabelFrame(master)
        self.master.frame_1.grid(row=0,column=0)
        self.title_list = ["Width(smallest)(mm)","Depth(mm)","P ultimate (kN)","Moment ultimate about width (2-2)(kNm)",\
                        "Moment ultimate about depth (3-3)(kNm)","fc'(N/mm2)","fy(N/mm2)","Reinforcement ratio(%)"]
        self.nrow = 0
        self.entry_set = defaultdict(list)
        for title in self.title_list:
            lbl = tk.Label(root.frame_1, text=title)
            lbl.grid(row=self.nrow, column=0, sticky='e')
            lbl.config(font=self.font_size) 
            ent = tk.Entry(root.frame_1)
            ent.grid(row=self.nrow, column=1)
            self.entry_set[title] = ent
            self.nrow += 1
        button = tk.Button(root.frame_1,text = "OK",width=8,relief = 'raised')
        button.bind('<Button-1>', self.assign)
        master.bind('<Return>', self.assign)
        button.grid(row = self.nrow,column=0,columnspan = 2,padx=10,pady=10)
    def assign(self,event):
        entry_dict = {k:float(v.get()) for k,v in self.entry_set.items()}
        #output
        self.width = entry_dict["Width(smallest)(mm)"]
        self.depth = entry_dict["Depth(mm)"]
        self.pu = entry_dict["P ultimate (kN)"]
        self.mu22 = entry_dict["Moment ultimate about width (2-2)(kNm)"]
        self.mu33 = entry_dict["Moment ultimate about depth (3-3)(kNm)"]
        self.ag = self.width * self.depth
        self.ast = self.ag * entry_dict["Reinforcement ratio(%)"]/100
        self.po = (0.85 * entry_dict["fc'(N/mm2)"] * (self.ag - self.ast) + entry_dict["fy(N/mm2)"] * self.ast)/1000
        i22 = self.inertia(self.mu22,self.width)
        i33 = self.inertia(self.mu33,self.depth)
        text_22 = "Modifier for weaker axis 2-2 is: " + str(self.modifier(i22))
        text_33 = "Modifier for stronger axis 3-3 is: " + str(self.modifier(i33))

        lbl_output22 = tk.Label(root.frame_1, text=text_22)
        lbl_output22.grid(row=self.nrow+1, column=0,columnspan = 2,sticky='w')
        lbl_output22.config(font=self.font_size) 

        lbl_output33 = tk.Label(root.frame_1, text=text_33)
        lbl_output33.grid(row=self.nrow+2, column=0,columnspan = 2,sticky='w')
        lbl_output33.config(font=self.font_size) 
    def modifier(self,i):
        """some thing has to be done to clear previous output if any"""
        if i < 0.35:
            return "0.50" # why text? to display number to two decimal places.
        elif i > 0.7:
            return "1.00"
        else:
            return round((i/0.7),2)
    def inertia(self,mu,depth):
        """As per 6.6.3.1.1 of ACI 318-14"""
        return (0.8 + 25 * self.ast/self.ag) * (1-mu*1000/(self.pu * depth) - 0.5 * self.pu/self.po)


if __name__ == '__main__':
    root = tk.Tk()
    inst_1  = App(root)
    root.mainloop()