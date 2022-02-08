import wx
import wx.xrc
import wx.dataview
import pandas as pd
import sys 
import zipfile
import numpy as np

luping = []
lupingA = []

class FeaturListPanel ( wx.Panel ):

    def __init__( self, parent ):
        wx.Panel.__init__ ( self, parent, id = wx.ID_ANY, pos = wx.DefaultPosition, size = wx.Size( 920,100 ), style = wx.TAB_TRAVERSAL )
    
        sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Periksa Berdasarkan :" ), wx.VERTICAL )
    
        self.m_scrolledWindow8 = wx.ScrolledWindow( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL|wx.VSCROLL )
        self.m_scrolledWindow8.SetScrollRate( 5, 5 )
        h_layout = wx.FlexGridSizer( 0, 5, 0, 0 )
        h_layout.SetFlexibleDirection( wx.BOTH )
        h_layout.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
           
                  
        if len(luping) != 0 :
            h_layout.Clear()
            for lup in range(len(luping)):
                self.m_checkBox8 = wx.CheckBox( self.m_scrolledWindow8, wx.ID_ANY, str(luping[lup]) , wx.DefaultPosition, wx.DefaultSize, 0 )
                h_layout.Add( self.m_checkBox8, 0, wx.ALL, 5 )
                self.m_checkBox8.Bind(wx.EVT_CHECKBOX, self.OnCheck)
            
        self.lis_featur = []    
        self.m_scrolledWindow8.SetSizer( h_layout )
        self.m_scrolledWindow8.Layout()
        h_layout.Fit( self.m_scrolledWindow8 )
        sbSizer3.Add( self.m_scrolledWindow8, 1, wx.EXPAND |wx.ALL, 5 )  

        self.indi = 0

        self.SetSizer( sbSizer3 )
        self.Layout()

    def __del__( self ):
        pass
    def OnCheck(self, e):
        cb = e.GetEventObject()
        if cb.GetValue() == True:
            if cb.GetLabel() not in self.lis_featur:
                self.lis_featur.append(cb.GetLabelText())
                
        else:
            self.lis_featur.remove(cb.GetLabelText())
            

            
    def ulang(self):
        if self.indi == 0:
            self.m_scrolledWindow8.Destroy()
            self.__init__(self)
        else :
            self.__init__(self)
             
        



class FeaturListPanelA ( wx.Panel ):

    def __init__( self, parent ):
        wx.Panel.__init__ ( self, parent, id = wx.ID_ANY, pos = wx.DefaultPosition, size = wx.Size( 920,100 ), style = wx.TAB_TRAVERSAL )
    
        sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Periksa Berdasarkan :" ), wx.VERTICAL )
    
        self.m_scrolledWindow8 = wx.ScrolledWindow( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL|wx.VSCROLL )
        self.m_scrolledWindow8.SetScrollRate( 5, 5 )
        h_layout = wx.FlexGridSizer( 0, 5, 0, 0 )
        h_layout.SetFlexibleDirection( wx.BOTH )
        h_layout.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
           
                  
        if len(lupingA) != 0 :
            h_layout.Clear()
            for lup in range(len(lupingA)):
                self.m_checkBox8 = wx.CheckBox( self.m_scrolledWindow8, wx.ID_ANY, str(lupingA[lup]) , wx.DefaultPosition, wx.DefaultSize, 0 )
                h_layout.Add( self.m_checkBox8, 0, wx.ALL, 5 )
                self.m_checkBox8.Bind(wx.EVT_CHECKBOX, self.OnCheck)
            
        self.lis_featur = []    
        self.m_scrolledWindow8.SetSizer( h_layout )
        self.m_scrolledWindow8.Layout()
        h_layout.Fit( self.m_scrolledWindow8 )
        sbSizer3.Add( self.m_scrolledWindow8, 1, wx.EXPAND |wx.ALL, 5 )  

        self.indi = 0

        self.SetSizer( sbSizer3 )
        self.Layout()

    def __del__( self ):
        pass
    def OnCheck(self, e):
        cb = e.GetEventObject()
        if cb.GetValue() == True:
            if cb.GetLabel() not in self.lis_featur:
                self.lis_featur.append(cb.GetLabelText())
                
        else:
            self.lis_featur.remove(cb.GetLabelText())
            

            
    def ulang(self):
        if self.indi == 0:
            self.m_scrolledWindow8.Destroy()
            self.__init__(self)
        else :
            self.__init__(self)
             
        




###########################################################################
## Class PanelDuplikasi
###########################################################################

class PanelDuplikasi ( wx.Panel ):

    def __init__( self, parent ):
        wx.Panel.__init__ ( self, parent, id = wx.ID_ANY, pos = wx.DefaultPosition, size = wx.Size( 1020,700 ), style = wx.TAB_TRAVERSAL )
    
        bSizer1 = wx.BoxSizer( wx.VERTICAL )
    
        self.m_scrolledWindow1 = wx.ScrolledWindow( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL|wx.VSCROLL )
        self.m_scrolledWindow1.SetScrollRate( 5, 5 )
        bSizer2 = wx.BoxSizer( wx.VERTICAL )
    
        self.m_staticText1 = wx.StaticText( self.m_scrolledWindow1, wx.ID_ANY, u"Upload file :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText1.Wrap( -1 )
        bSizer2.Add( self.m_staticText1, 0, wx.ALL, 5 )
    
        self.m_uploadfile = wx.FilePickerCtrl( self.m_scrolledWindow1, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.xlsx*", wx.DefaultPosition, wx.Size( -1,30 ), wx.FLP_DEFAULT_STYLE )
        bSizer2.Add( self.m_uploadfile, 0, wx.ALL|wx.EXPAND, 5 )
    
        fgSizer1 = wx.FlexGridSizer( 0, 3, 0, 0 )
        fgSizer1.SetFlexibleDirection( wx.BOTH )
        fgSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
    
        self.m_staticText2 = wx.StaticText( self.m_scrolledWindow1, wx.ID_ANY, u"Sheet yang Digunakan :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText2.Wrap( -1 )
        fgSizer1.Add( self.m_staticText2, 0, wx.ALL, 5 )
    
        self.m_staticText3 = wx.StaticText( self.m_scrolledWindow1, wx.ID_ANY, u"Header :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText3.Wrap( -1 )
        fgSizer1.Add( self.m_staticText3, 0, wx.ALL, 5 )
    
        self.m_staticText4 = wx.StaticText( self.m_scrolledWindow1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText4.Wrap( -1 )
        fgSizer1.Add( self.m_staticText4, 0, wx.ALL, 5 )
    
        self.m_sheetname = wx.TextCtrl( self.m_scrolledWindow1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 300,-1 ), 0 )
        fgSizer1.Add( self.m_sheetname, 0, wx.ALL, 5 )
    
        self.m_headerset = wx.SpinCtrl( self.m_scrolledWindow1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 100,-1 ), wx.SP_ARROW_KEYS, 1, 999, 1 )
        fgSizer1.Add( self.m_headerset, 0, wx.ALL, 5 )
    
        self.btn_submit = wx.Button( self.m_scrolledWindow1, wx.ID_ANY, u"Submit", wx.DefaultPosition, wx.Size( 100,-1 ), 0 )
        fgSizer1.Add( self.btn_submit, 0, wx.ALL, 5 )
        self.btn_submit.Bind(wx.EVT_BUTTON, self.OnSubmit)
    
    
        bSizer2.Add( fgSizer1, 0, wx.EXPAND, 5 )
    
        self.m_panel7 = wx.Panel( self.m_scrolledWindow1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        v_layout = wx.BoxSizer( wx.VERTICAL )
        
        self.daftarfeature = FeaturListPanel(self.m_panel7)
        v_layout.Add(self.daftarfeature, 0, wx.EXPAND |wx.ALL, 5)
    
    
        self.m_panel7.SetSizer( v_layout )
        self.m_panel7.Layout()
        v_layout.Fit( self.m_panel7 )
        bSizer2.Add( self.m_panel7, 0, wx.EXPAND |wx.ALL, 5 )
    
        fgSizer3 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer3.SetFlexibleDirection( wx.BOTH )
        fgSizer3.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
    
        self.btn_periksa = wx.Button( self.m_scrolledWindow1, wx.ID_ANY, u"Periksa", wx.DefaultPosition, wx.Size( 130,30 ), 0 )
        fgSizer3.Add( self.btn_periksa, 0, wx.ALL, 5 )
        self.btn_periksa.Bind(wx.EVT_BUTTON, self.OnPeriksa)
        self.btn_periksa.Disable()
    
        self.btn_reset = wx.Button( self.m_scrolledWindow1, wx.ID_ANY, u"Reset", wx.DefaultPosition, wx.Size( 100,30 ), 0 )
        fgSizer3.Add( self.btn_reset, 0, wx.ALL, 5 )
        self.btn_reset.Bind(wx.EVT_BUTTON, self.OnReset)
        self.btn_reset.Disable()
    
    
        bSizer2.Add( fgSizer3, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
    
        self.m_datahasil = wx.dataview.DataViewListCtrl( self.m_scrolledWindow1, wx.ID_ANY, wx.DefaultPosition, wx.Size( -1,250 ), wx.dataview.DV_HORIZ_RULES|wx.dataview.DV_ROW_LINES )
        bSizer2.Add( self.m_datahasil, 0, wx.ALL|wx.EXPAND, 5 )
    
        self.btn_unduh = wx.Button( self.m_scrolledWindow1, wx.ID_ANY, u"Unduh File", wx.DefaultPosition, wx.Size( 200,30 ), 0 )
        bSizer2.Add( self.btn_unduh, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
        self.btn_unduh.Bind(wx.EVT_BUTTON, self.OnUnduh)
        self.btn_unduh.Disable()
    
    
        self.m_scrolledWindow1.SetSizer( bSizer2 )
        self.m_scrolledWindow1.Layout()
        bSizer2.Fit( self.m_scrolledWindow1 )
        bSizer1.Add( self.m_scrolledWindow1, 1, wx.EXPAND |wx.ALL, 5 )
        
        #self.luping = []
    
    
        self.SetSizer( bSizer1 )
        self.Layout()
        
    
    def __del__( self ):
        pass

    def OnSubmit(self, e): 
        luping.clear()
        path = self.m_uploadfile.GetPath()
        self.btn_periksa.Enable()
        self.btn_reset.Enable()
        self.m_datahasil.DeleteAllItems()
        self.m_datahasil.ClearColumns()        
        try:
            if self.m_sheetname.IsEmpty():
                df = pd.read_excel(path, engine='openpyxl', header=int(self.m_headerset.GetValue()-1))
                kol = list(df)
                
                for x in range(len(kol)):    
                    luping.append(kol[x]) 
            else:
                df = pd.read_excel(path,sheet_name=str(self.m_sheetname.GetValue()), engine='openpyxl', header=int(self.m_headerset.GetValue()-1))
                kol = list(df)
                
                for x in range(len(kol)):    
                    luping.append(kol[x])                
            
            self.daftarfeature.ulang()
            #self.luping.clear()
            
        except FileNotFoundError :
            msg = wx.MessageDialog(self, 'Anda belum upload file', 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy() 
            self.btn_periksa.Disable()
            self.btn_reset.Disable() 
        
        except ValueError as err :
            pesan = 'Terjadi kesalahan!!\n\nError : '+str(err)+'\n\nPeriksa dan sesuaikan kembali inputan anda'
            msg = wx.MessageDialog(self, pesan, 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            
        except:
            msg = wx.MessageDialog(self, 'Mohon maaf, Terjadi kesalahan\nJalankan ulang', 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            self.btn_unduh.Disable()        



    def OnPeriksa(self, e):
        feature = self.daftarfeature.lis_featur
        if len(feature) == 0:
            msg = wx.MessageDialog(self, 'Centang salah satu pilihan untuk hasil yang lebih akurat', 'Information', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            self.m_datahasil.DeleteAllItems()
            self.m_datahasil.ClearColumns()            
        else:
            path = self.m_uploadfile.GetPath()
            self.m_datahasil.DeleteAllItems()
            self.m_datahasil.ClearColumns()                            
           
            try:
                if self.m_sheetname.IsEmpty():
                    df = pd.read_excel(path,engine='openpyxl', header=int(self.m_headerset.GetValue()-1))
                    self.df_result = df_result = df[df[feature].duplicated()]
                    self.nama_kolom = list(df_result)
                    
                    for kol in range(len(self.nama_kolom)):    
                        self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                    
                    for row in range(len(df_result)):
                        lis_temp = list(df_result.iloc[row])
                        dup = []
                        for x in range(len(lis_temp)):
                            dup.append(str(lis_temp[x]))
                
                        self.m_datahasil.AppendItem(dup)  
                        dup.clear()
                        
                    self.btn_unduh.Enable()
                
                #================================================================================================================
                
                else:
                    try :
                        df = pd.read_excel(path,sheet_name=str(self.m_sheetname.GetValue()),engine='openpyxl', header=int(self.m_headerset.GetValue()-1))
                        self.df_result = df_result = df[df[feature].duplicated()]
                        self.nama_kolom = list(df_result)
                        
                        for kol in range(len(self.nama_kolom)):    
                            self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                    
                        for row in range(len(df_result)):
                            lis_temp = list(df_result.iloc[row])
                            dup = []
                            for x in range(len(lis_temp)):
                                dup.append(str(lis_temp[x]))
                    
                            self.m_datahasil.AppendItem(dup)  
                            dup.clear()
                        
                        self.btn_unduh.Enable()
                            
                    except ValueError :
                        msg = wx.MessageDialog(self, 'Terjadi kesalahan!!\nPeriksa dan sesuaikan kembali inputan anda', 'Peringatan', wx.OK_DEFAULT)
                        ans = msg.ShowModal()
                        msg.Destroy() 
                        self.btn_unduh.Disable()
            
            except FileNotFoundError :
                msg = wx.MessageDialog(self, 'Anda belum upload file', 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()
                self.btn_unduh.Disable()
                            
            except ValueError as err :
                pesan = 'Terjadi kesalahan!!\n\nError : '+str(err)+'\n\nPeriksa dan sesuaikan kembali inputan anda'
                msg = wx.MessageDialog(self, pesan, 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()
                
            except:
                msg = wx.MessageDialog(self, 'Mohon maaf, Terjadi kesalahan\nJalankan ulang', 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()
                self.btn_unduh.Disable()
            
    
    def OnReset(self, e):
        self.m_datahasil.DeleteAllItems()
        self.m_datahasil.ClearColumns()
        self.daftarfeature.m_scrolledWindow8.Destroy()
        self.daftarfeature.indi = 1
        self.m_uploadfile.SetPath("")
        self.m_sheetname.SetLabel("")
        self.m_headerset.SetValue(1)
        self.btn_periksa.Disable()
        self.btn_reset.Disable()
        self.btn_unduh.Disable()
    
    def onSave(self, e):
        save = []
        for row in range(len(self.df_result)):
            lis_temp = list(self.df_result.iloc[row])
            dup = []
            for x in range(len(lis_temp)):
                temp = "'"+str(lis_temp[x])
                dup.append(temp)
            save.append(dup)
        m_save = pd.DataFrame(save, columns=self.nama_kolom)
        
        return m_save.to_excel(e)
    
    def OnUnduh(self, event):
    
        with wx.FileDialog(self, "Save Excel file", wildcard="Excel files (*.xlsx)|*.xlsx",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
    
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                fileDialog.Close()
                
            else:
                pathname = fileDialog.GetPath()
                try:            
                    with open(pathname, 'wb') as file:
                        self.onSave(file)
                    msg = wx.MessageDialog(self, 'Berhasil disimpan', 'Information', wx.OK_DEFAULT)
                    ans = msg.ShowModal()
                    msg.Destroy()                    
                except IOError:
                    wx.LogError("Tidak dapat menyimpan data saat ini dalam file '%s'." % pathname)    




###########################################################################
## Class PanelCompare
###########################################################################

class PanelCompare ( wx.Panel ):

    def __init__( self, parent ):
        wx.Panel.__init__ ( self, parent, id = wx.ID_ANY, pos = wx.DefaultPosition, size = wx.Size( 1000,700 ), style = wx.TAB_TRAVERSAL )
    
        bSizer6 = wx.BoxSizer( wx.VERTICAL )
    
        self.m_scrolledWindow3 = wx.ScrolledWindow( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL|wx.VSCROLL )
        self.m_scrolledWindow3.SetScrollRate( 5, 5 )
        layout_utama = wx.BoxSizer( wx.VERTICAL )
    
        self.m_staticText5 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, u"Upload File 1 :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText5.Wrap( -1 )
        layout_utama.Add( self.m_staticText5, 0, wx.ALL, 5 )
    
        self.m_uploadfileA = wx.FilePickerCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.xlsx*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
        layout_utama.Add( self.m_uploadfileA, 0, wx.ALL|wx.EXPAND, 5 )
    
        fgSizer4 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer4.SetFlexibleDirection( wx.BOTH )
        fgSizer4.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
    
        self.m_staticText6 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, u"Sheet yang Digunakan :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText6.Wrap( -1 )
        fgSizer4.Add( self.m_staticText6, 0, wx.ALL, 5 )
    
        self.m_staticText7 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, u"Header", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText7.Wrap( -1 )
        fgSizer4.Add( self.m_staticText7, 0, wx.ALL, 5 )
    
        self.m_sheetnameA = wx.TextCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 300,-1 ), 0 )
        fgSizer4.Add( self.m_sheetnameA, 0, wx.ALL, 5 )
    
        self.m_headersetA = wx.SpinCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 100,-1 ), wx.SP_ARROW_KEYS, 1, 999, 1 )
        fgSizer4.Add( self.m_headersetA, 0, wx.ALL, 5 )
    
    
        layout_utama.Add( fgSizer4, 1, wx.EXPAND, 5 )
    
        self.pnl_featurlistA = wx.Panel( self.m_scrolledWindow3, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        v_layoutA = wx.BoxSizer( wx.VERTICAL )
        
        self.daftarfeatureA = FeaturListPanelA(self.pnl_featurlistA)
        v_layoutA.Add(self.daftarfeatureA, 0, wx.ALL|wx.EXPAND, 5)
    
    
        self.pnl_featurlistA.SetSizer( v_layoutA )
        self.pnl_featurlistA.Layout()
        v_layoutA.Fit( self.pnl_featurlistA )
        layout_utama.Add( self.pnl_featurlistA, 0, wx.ALL|wx.EXPAND, 5 )
    
    
        #layout_utama.AddSpacer( ( 0, 20), 0, 0, 5 )
    
        self.m_staticText8 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, u"Upload File 2 :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText8.Wrap( -1 )
        layout_utama.Add( self.m_staticText8, 0, wx.ALL, 5 )
    
        self.m_uploadfileB = wx.FilePickerCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.xlsx*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
        layout_utama.Add( self.m_uploadfileB, 0, wx.ALL|wx.EXPAND, 5 )
    
        self.m_checkBox1 = wx.CheckBox( self.m_scrolledWindow3, wx.ID_ANY, u"Gunakan File yang Sama", wx.DefaultPosition, wx.DefaultSize, 0 )
        layout_utama.Add( self.m_checkBox1, 0, wx.ALL, 5 )
        self.m_checkBox1.Bind(wx.EVT_CHECKBOX, self.OnCheck)
    
        fgSizer5 = wx.FlexGridSizer( 0, 3, 0, 0 )
        fgSizer5.SetFlexibleDirection( wx.BOTH )
        fgSizer5.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
    
        self.m_staticText9 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, u"Sheet yang Digunakan :", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText9.Wrap( -1 )
        fgSizer5.Add( self.m_staticText9, 0, wx.ALL, 5 )
    
        self.m_staticText10 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, u"Header", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText10.Wrap( -1 )
        fgSizer5.Add( self.m_staticText10, 0, wx.ALL, 5 )
    
        self.m_staticText11 = wx.StaticText( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText11.Wrap( -1 )
        fgSizer5.Add( self.m_staticText11, 0, wx.ALL, 5 )
    
        self.m_sheetnameB = wx.TextCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 300,-1 ), 0 )
        fgSizer5.Add( self.m_sheetnameB, 0, wx.ALL, 5 )
    
        self.m_headersetB = wx.SpinCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 100,-1 ), wx.SP_ARROW_KEYS, 1, 999, 1 )
        fgSizer5.Add( self.m_headersetB, 0, wx.ALL, 5 )
    
        self.btn_submit = wx.Button( self.m_scrolledWindow3, wx.ID_ANY, u"Submit", wx.DefaultPosition, wx.Size( 100,-1 ), 0 )
        fgSizer5.Add( self.btn_submit, 0, wx.ALL, 5 )
        self.btn_submit.Bind(wx.EVT_BUTTON, self.OnSubmit)
    
    
        layout_utama.Add( fgSizer5, 1, wx.EXPAND, 5 )
    
        self.pnl_featurlist = wx.Panel( self.m_scrolledWindow3, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        v_layout = wx.BoxSizer( wx.VERTICAL )
        
        self.daftarfeature = FeaturListPanel(self.pnl_featurlist)
        v_layout.Add(self.daftarfeature, 0, wx.EXPAND |wx.ALL, 5)
    
    
        self.pnl_featurlist.SetSizer( v_layout )
        self.pnl_featurlist.Layout()
        v_layout.Fit( self.pnl_featurlist )
        layout_utama.Add( self.pnl_featurlist, 0, wx.EXPAND |wx.ALL, 5 )
    
        fgSizer6 = wx.FlexGridSizer( 0, 3, 0, 0 )
        fgSizer6.SetFlexibleDirection( wx.BOTH )
        fgSizer6.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
    
        self.btn_periksa = wx.Button( self.m_scrolledWindow3, wx.ID_ANY, u"Cek Yang Ada", wx.DefaultPosition, wx.Size( 130,30 ), 0 )
        fgSizer6.Add( self.btn_periksa, 0, wx.ALL, 5 )
        self.btn_periksa.Bind(wx.EVT_BUTTON, self.OnPeriksa)
        self.btn_periksa.Disable()
        
        self.btn_tdkada = wx.Button( self.m_scrolledWindow3, wx.ID_ANY, u"Cek Yang Tidak Ada", wx.DefaultPosition, wx.Size( 130,30 ), 0 )
        fgSizer6.Add( self.btn_tdkada, 0, wx.ALL, 5 )
        self.btn_tdkada.Bind(wx.EVT_BUTTON, self.OnPeriksaTidakAda)
        self.btn_tdkada.Disable()
    
        self.btn_reset = wx.Button( self.m_scrolledWindow3, wx.ID_ANY, u"Reset", wx.DefaultPosition, wx.Size( 100,30 ), 0 )
        fgSizer6.Add( self.btn_reset, 0, wx.ALL, 5 )
        self.btn_reset.Bind(wx.EVT_BUTTON, self.OnReset)
        self.btn_reset.Disable()
    
    
        layout_utama.Add( fgSizer6, 1, wx.ALIGN_RIGHT|wx.ALL, 5 )
    
        self.m_datahasil = wx.dataview.DataViewListCtrl( self.m_scrolledWindow3, wx.ID_ANY, wx.DefaultPosition, wx.Size( -1,250 ), wx.dataview.DV_HORIZ_RULES|wx.dataview.DV_ROW_LINES )
        layout_utama.Add( self.m_datahasil, 0, wx.ALL|wx.EXPAND, 5 )
    
        self.btn_unduh = wx.Button( self.m_scrolledWindow3, wx.ID_ANY, u"Unduh File", wx.DefaultPosition, wx.Size( 200,30 ), 0 )
        layout_utama.Add( self.btn_unduh, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
        self.btn_unduh.Bind(wx.EVT_BUTTON, self.OnUnduh)
        self.btn_unduh.Disable()
    
    
        self.m_scrolledWindow3.SetSizer( layout_utama )
        self.m_scrolledWindow3.Layout()
        layout_utama.Fit( self.m_scrolledWindow3 )
        bSizer6.Add( self.m_scrolledWindow3, 1, wx.EXPAND |wx.ALL, 5 )
    
    
        self.SetSizer( bSizer6 )
        self.Layout()

    def __del__( self ):
        pass
    
    def OnPeriksa (self, e):
        feature = self.daftarfeature.lis_featur
        featureA = self.daftarfeatureA.lis_featur
        if len(feature) == 0 or len(featureA) == 0:
            msg = wx.MessageDialog(self, 'Centang salah satu pilihan "Periksa Berdasarkan" untuk hasil yang lebih akurat', 'Information', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            self.m_datahasil.DeleteAllItems()
            self.m_datahasil.ClearColumns()
        else :
            pathA = self.m_uploadfileA.GetPath()
            pathB = self.m_uploadfileB.GetPath()
            self.m_datahasil.DeleteAllItems()
            self.m_datahasil.ClearColumns()
            
            try:
                if self.m_sheetnameA.IsEmpty():
                    dfA = pd.read_excel(pathA, engine='openpyxl', header=int(self.m_headersetA.GetValue()-1))
                    if self.m_sheetnameB.IsEmpty():
                        dfB = pd.read_excel(pathB, engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                        if dfA.dtypes[featureA[0]] == np.float:
                            dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                            
                        cekA = np.array(dfA[featureA])
                        cekB = np.array(dfB[feature]) 
                        hasil = []
                        for lupA in range(len(cekA)):
                            if cekA[lupA] in cekB:
                                samadengan = pathB.split('\\')
                                ket = 'ADA JUGA DI "'+str(samadengan[-1])+'"'
                                hasil.append(ket)
                            else :
                                hasil.append(' ')
                        
                        dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                        
                        
                        self.df_result = pd.DataFrame()
                        self.nama_kolom = list(dfA)      
                        for kol in range(len(self.nama_kolom)):    
                            self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                            
                        self.save = []
                        for lupA in range(len(cekA)):
                            lis_temp = list(dfA.iloc[lupA])
                            found = []
                            self.dup_save = []
                            for x in range(len(lis_temp)):
                                temp = "'"+str(lis_temp[x])
                                found.append(str(lis_temp[x]))
                                self.dup_save.append(temp)
                                
                            self.save.append(self.dup_save)
                            self.m_datahasil.AppendItem(found)
                            found.clear()                                    
                        
                        
                        self.btn_unduh.Enable()
                                                                                                          
                    
                    else :
                        sheetB = self.m_sheetnameB.GetValue()
                        sheetB = sheetB.split(", ")
                        
                        if len(sheetB) > 1 :
                            for sheet in list(sheetB):
                                dfB = pd.read_excel(pathB,sheet_name=str(sheet) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                                cekA = np.array(dfA[featureA])
                                cekB = np.array(dfB[feature]) 
                                hasil = []
                                for lupA in range(len(cekA)):
                                    if cekA[lupA] in cekB:
                                        hasil.append('Ada')
                                    else :
                                        hasil.append(' ')
                                
                                fitur = "DI DATA "+str(sheet)
                                dfA[str(fitur)] = hasil
                            
                            self.df_result = pd.DataFrame()
                            self.nama_kolom = list(dfA)      
                            for kol in range(len(self.nama_kolom)):    
                                self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                            
                                
                            self.save = []
                            for lupA in range(len(dfA)):
                                lis_temp = list(dfA.iloc[lupA])
                                found = []
                                self.dup_save = []
                                for x in range(len(lis_temp)):
                                    temp = "'"+str(lis_temp[x])
                                    found.append(str(lis_temp[x]))
                                    self.dup_save.append(temp)
                                    
                                self.save.append(self.dup_save)
                                self.m_datahasil.AppendItem(found)
                                found.clear()                                    
                            
                                
                        
                        
                        else :
                            dfB = pd.read_excel(pathB,sheet_name=str(self.m_sheetnameB.GetValue()) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                            if dfA.dtypes[featureA[0]] == np.float:
                                dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                                
                            cekA = np.array(dfA[featureA])
                            cekB = np.array(dfB[feature]) 
                            hasil = []
                            for lupA in range(len(cekA)):
                                if cekA[lupA] in cekB:
                                    samadengan = pathB.split('\\')
                                    ket = 'ADA JUGA DI "'+str(samadengan[-1])+'" PADA SHEET "'+str(self.m_sheetnameB.GetValue())+'"'
                                    hasil.append(ket)
                                else :
                                    hasil.append(' ')
                            
                            dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                            
                            
                            self.df_result = pd.DataFrame()
                            self.nama_kolom = list(dfA)      
                            for kol in range(len(self.nama_kolom)):    
                                self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                
                            self.save = []
                            for lupA in range(len(cekA)):
                                lis_temp = list(dfA.iloc[lupA])
                                found = []
                                self.dup_save = []
                                for x in range(len(lis_temp)):
                                    temp = "'"+str(lis_temp[x])
                                    found.append(str(lis_temp[x]))
                                    self.dup_save.append(temp)
                                    
                                self.save.append(self.dup_save)
                                self.m_datahasil.AppendItem(found)
                                found.clear()                                    
                            
                            
                            self.btn_unduh.Enable()                                
                                                
                                   
                else :
                    sheetA = self.m_sheetnameA.GetValue()
                    sheetA = sheetA.split(", ")
                    
                    
                    if len(sheetA) > 1 :
                        msg = wx.MessageDialog(self, 'Sheet file 1 tidak dapat diisi lebih dari satu sheet', 'Information', wx.OK_DEFAULT)
                        ans = msg.ShowModal()
                        msg.Destroy()
                        self.m_datahasil.DeleteAllItems()
                        self.m_datahasil.ClearColumns()
                        
                    else :
                        dfA = pd.read_excel(pathA,sheet_name=str(self.m_sheetnameA.GetValue()) ,engine='openpyxl', header=int(self.m_headersetA.GetValue()-1))
                        
                        if self.m_sheetnameB.IsEmpty():
                            dfB = pd.read_excel(pathB, engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                            if dfA.dtypes[featureA[0]] == np.float:
                                dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                                
                            cekA = np.array(dfA[featureA])
                            cekB = np.array(dfB[feature]) 
                            hasil = []
                            for lupA in range(len(cekA)):
                                if cekA[lupA] in cekB:
                                    samadengan = pathB.split('\\')
                                    ket = 'ADA JUGA DI "'+str(samadengan[-1])+'"'
                                    hasil.append(ket)
                                else :
                                    hasil.append(' ')
                            
                            dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                            
                            
                            self.df_result = pd.DataFrame()
                            self.nama_kolom = list(dfA)      
                            for kol in range(len(self.nama_kolom)):    
                                self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                
                            self.save = []
                            for lupA in range(len(cekA)):
                                lis_temp = list(dfA.iloc[lupA])
                                found = []
                                self.dup_save = []
                                for x in range(len(lis_temp)):
                                    temp = "'"+str(lis_temp[x])
                                    found.append(str(lis_temp[x]))
                                    self.dup_save.append(temp)
                                    
                                self.save.append(self.dup_save)
                                self.m_datahasil.AppendItem(found)
                                found.clear()                                    
                            
                            
                            self.btn_unduh.Enable()
                                                                                                              
                        
                        else :
                            sheetB = self.m_sheetnameB.GetValue()
                            sheetB = sheetB.split(", ")
                            
                            if len(sheetB) > 1 :
                                for sheet in list(sheetB):
                                    dfB = pd.read_excel(pathB,sheet_name=str(sheet) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                                    cekA = np.array(dfA[featureA])
                                    cekB = np.array(dfB[feature])
                                    hasil = []
                                    for lupA in range(len(cekA)):
                                        if cekA[lupA] in cekB:
                                            hasil.append('Ada')
                                            
                                        else :
                                            hasil.append(' ')
                                    
                                    fitur = "DI DATA "+str(sheet)
                                    dfA[str(fitur)] = hasil
                                
                                self.df_result = pd.DataFrame()
                                self.nama_kolom = list(dfA)      
                                for kol in range(len(self.nama_kolom)):    
                                    self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                
                                    
                                self.save = []
                                for lupA in range(len(dfA)):
                                    lis_temp = list(dfA.iloc[lupA])
                                    found = []
                                    self.dup_save = []
                                    for x in range(len(lis_temp)):
                                        temp = "'"+str(lis_temp[x])
                                        found.append(str(lis_temp[x]))
                                        self.dup_save.append(temp)
                                        
                                    self.save.append(self.dup_save)
                                    self.m_datahasil.AppendItem(found)
                                    found.clear()
                                
                                self.btn_unduh.Enable()
                                
                                    
                            
                            
                            else :
                                dfB = pd.read_excel(pathB,sheet_name=str(self.m_sheetnameB.GetValue()) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                                if dfA.dtypes[featureA[0]] == np.float:
                                    dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                                    
                                cekA = np.array(dfA[featureA])
                                cekB = np.array(dfB[feature])
                                hasil = []
                                for lupA in range(len(cekA)):
                                    if cekA[lupA] in cekB:
                                        samadengan = pathB.split('\\')
                                        ket = 'ADA JUGA DI "'+str(samadengan[-1])+'" PADA SHEET "'+str(self.m_sheetnameB.GetValue())+'"'
                                        hasil.append(ket)
                                    else :
                                        hasil.append(' ')
                                
                                dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                                
                                
                                self.df_result = pd.DataFrame()
                                self.nama_kolom = list(dfA)      
                                for kol in range(len(self.nama_kolom)):    
                                    self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                    
                                self.save = []
                                for lupA in range(len(cekA)):
                                    lis_temp = list(dfA.iloc[lupA])
                                    found = []
                                    self.dup_save = []
                                    for x in range(len(lis_temp)):
                                        temp = "'"+str(lis_temp[x])
                                        found.append(str(lis_temp[x]))
                                        self.dup_save.append(temp)
                                        
                                    self.save.append(self.dup_save)
                                    self.m_datahasil.AppendItem(found)
                                    found.clear()                                    
                                
                                
                                self.btn_unduh.Enable()                                
                                
                            
            except FileNotFoundError :
                msg = wx.MessageDialog(self, 'Anda belum upload file', 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy() 
                
            
            except ValueError as err:
                pesan = 'Terjadi kesalahan!!\n\nError : '+str(err)+'\n\nPeriksa dan sesuaikan kembali inputan anda'
                msg = wx.MessageDialog(self, pesan, 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()           
                
                
            except:
                msg = wx.MessageDialog(self, 'Mohon maaf, Terjadi kesalahan\nJalankan ulang', 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()
                          
            
            
      
    def OnPeriksaTidakAda (self, e):
        feature = self.daftarfeature.lis_featur
        featureA = self.daftarfeatureA.lis_featur
        if len(feature) == 0 or len(featureA) == 0:
            msg = wx.MessageDialog(self, 'Centang salah satu pilihan "Periksa Berdasarkan" untuk hasil yang lebih akurat', 'Information', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            self.m_datahasil.DeleteAllItems()
            self.m_datahasil.ClearColumns()
        else :
            pathA = self.m_uploadfileA.GetPath()
            pathB = self.m_uploadfileB.GetPath()
            self.m_datahasil.DeleteAllItems()
            self.m_datahasil.ClearColumns()
            
            try:
                if self.m_sheetnameA.IsEmpty():
                    dfA = pd.read_excel(pathA, engine='openpyxl', header=int(self.m_headersetA.GetValue()-1))
                    if self.m_sheetnameB.IsEmpty():
                        dfB = pd.read_excel(pathB, engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                        if dfA.dtypes[featureA[0]] == np.float:
                            dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                            
                        cekA = np.array(dfA[featureA])
                        cekB = np.array(dfB[feature]) 
                        hasil = []
                        for lupA in range(len(cekA)):
                            if cekA[lupA] not in cekB:
                                samadengan = pathB.split('\\')
                                ket = 'TIDAK ADA DI "'+str(samadengan[-1])+'"'
                                hasil.append(ket)
                            else :
                                hasil.append(' ')
                        
                        dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                        
                        
                        self.df_result = pd.DataFrame()
                        self.nama_kolom = list(dfA)      
                        for kol in range(len(self.nama_kolom)):    
                            self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                            
                        self.save = []
                        for lupA in range(len(cekA)):
                            lis_temp = list(dfA.iloc[lupA])
                            found = []
                            self.dup_save = []
                            for x in range(len(lis_temp)):
                                temp = "'"+str(lis_temp[x])
                                found.append(str(lis_temp[x]))
                                self.dup_save.append(temp)
                                
                            self.save.append(self.dup_save)
                            self.m_datahasil.AppendItem(found)
                            found.clear()                                    
                        
                        
                        self.btn_unduh.Enable()
                                                                                                          
                    
                    else :
                        sheetB = self.m_sheetnameB.GetValue()
                        sheetB = sheetB.split(", ")
                        
                        if len(sheetB) > 1 :
                            for sheet in list(sheetB):
                                dfB = pd.read_excel(pathB,sheet_name=str(sheet) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                                cekA = np.array(dfA[featureA])
                                cekB = np.array(dfB[feature]) 
                                hasil = []
                                for lupA in range(len(cekA)):
                                    if cekA[lupA] not in cekB:
                                        hasil.append('Tidak Ada')
                                    else :
                                        hasil.append(' ')
                                
                                fitur = "DI DATA "+str(sheet)
                                dfA[str(fitur)] = hasil
                            
                            self.df_result = pd.DataFrame()
                            self.nama_kolom = list(dfA)      
                            for kol in range(len(self.nama_kolom)):    
                                self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                            
                                
                            self.save = []
                            for lupA in range(len(dfA)):
                                lis_temp = list(dfA.iloc[lupA])
                                found = []
                                self.dup_save = []
                                for x in range(len(lis_temp)):
                                    temp = "'"+str(lis_temp[x])
                                    found.append(str(lis_temp[x]))
                                    self.dup_save.append(temp)
                                    
                                self.save.append(self.dup_save)
                                self.m_datahasil.AppendItem(found)
                                found.clear()                                    
                            
                                
                        
                        
                        else :
                            dfB = pd.read_excel(pathB,sheet_name=str(self.m_sheetnameB.GetValue()) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                            if dfA.dtypes[featureA[0]] == np.float:
                                dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                                
                            cekA = np.array(dfA[featureA])
                            cekB = np.array(dfB[feature]) 
                            hasil = []
                            for lupA in range(len(cekA)):
                                if cekA[lupA] not in cekB:
                                    samadengan = pathB.split('\\')
                                    ket = 'TIDAK ADA DI "'+str(samadengan[-1])+'" PADA SHEET "'+str(self.m_sheetnameB.GetValue())+'"'
                                    hasil.append(ket)
                                else :
                                    hasil.append(' ')
                            
                            dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                            
                            
                            self.df_result = pd.DataFrame()
                            self.nama_kolom = list(dfA)      
                            for kol in range(len(self.nama_kolom)):    
                                self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                
                            self.save = []
                            for lupA in range(len(cekA)):
                                lis_temp = list(dfA.iloc[lupA])
                                found = []
                                self.dup_save = []
                                for x in range(len(lis_temp)):
                                    temp = "'"+str(lis_temp[x])
                                    found.append(str(lis_temp[x]))
                                    self.dup_save.append(temp)
                                    
                                self.save.append(self.dup_save)
                                self.m_datahasil.AppendItem(found)
                                found.clear()                                    
                            
                            
                            self.btn_unduh.Enable()                                
                                                
                                   
                else :
                    sheetA = self.m_sheetnameA.GetValue()
                    sheetA = sheetA.split(", ")
                    
                    
                    if len(sheetA) > 1 :
                        msg = wx.MessageDialog(self, 'Sheet file 1 tidak dapat diisi lebih dari satu sheet', 'Information', wx.OK_DEFAULT)
                        ans = msg.ShowModal()
                        msg.Destroy()
                        self.m_datahasil.DeleteAllItems()
                        self.m_datahasil.ClearColumns()
                        
                    else :
                        dfA = pd.read_excel(pathA,sheet_name=str(self.m_sheetnameA.GetValue()) ,engine='openpyxl', header=int(self.m_headersetA.GetValue()-1))
                        
                        if self.m_sheetnameB.IsEmpty():
                            dfB = pd.read_excel(pathB, engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                            if dfA.dtypes[featureA[0]] == np.float:
                                dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                                
                            cekA = np.array(dfA[featureA])
                            cekB = np.array(dfB[feature]) 
                            hasil = []
                            for lupA in range(len(cekA)):
                                if cekA[lupA] not in cekB:
                                    samadengan = pathB.split('\\')
                                    ket = 'TIDAK ADA DI "'+str(samadengan[-1])+'"'
                                    hasil.append(ket)
                                else :
                                    hasil.append(' ')
                            
                            dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                            
                            
                            self.df_result = pd.DataFrame()
                            self.nama_kolom = list(dfA)      
                            for kol in range(len(self.nama_kolom)):    
                                self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                
                            self.save = []
                            for lupA in range(len(cekA)):
                                lis_temp = list(dfA.iloc[lupA])
                                found = []
                                self.dup_save = []
                                for x in range(len(lis_temp)):
                                    temp = "'"+str(lis_temp[x])
                                    found.append(str(lis_temp[x]))
                                    self.dup_save.append(temp)
                                    
                                self.save.append(self.dup_save)
                                self.m_datahasil.AppendItem(found)
                                found.clear()                                    
                            
                            
                            self.btn_unduh.Enable()
                                                                                                              
                        
                        else :
                            sheetB = self.m_sheetnameB.GetValue()
                            sheetB = sheetB.split(", ")
                            
                            if len(sheetB) > 1 :
                                for sheet in list(sheetB):
                                    dfB = pd.read_excel(pathB,sheet_name=str(sheet) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                                    cekA = np.array(dfA[featureA])
                                    cekB = np.array(dfB[feature]) 
                                    hasil = []
                                    for lupA in range(len(cekA)):
                                        if cekA[lupA] not in cekB:
                                            hasil.append('Tidak Ada')
                                        else :
                                            hasil.append(' ')
                                    
                                    fitur = "DI DATA "+str(sheet)
                                    dfA[str(fitur)] = hasil
                                
                                self.df_result = pd.DataFrame()
                                self.nama_kolom = list(dfA)      
                                for kol in range(len(self.nama_kolom)):    
                                    self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                
                                    
                                self.save = []
                                for lupA in range(len(dfA)):
                                    lis_temp = list(dfA.iloc[lupA])
                                    found = []
                                    self.dup_save = []
                                    for x in range(len(lis_temp)):
                                        temp = "'"+str(lis_temp[x])
                                        found.append(str(lis_temp[x]))
                                        self.dup_save.append(temp)
                                        
                                    self.save.append(self.dup_save)
                                    self.m_datahasil.AppendItem(found)
                                    found.clear()                                    
                                
                                self.btn_unduh.Enable()
                            
                            
                            else :
                                dfB = pd.read_excel(pathB,sheet_name=str(self.m_sheetnameB.GetValue()) ,engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                                if dfA.dtypes[featureA[0]] == np.float:
                                    dfA[featureA[0]] = dfA[featureA[0]].astype('int64')
                                    
                                cekA = np.array(dfA[featureA])
                                cekB = np.array(dfB[feature]) 
                                hasil = []
                                for lupA in range(len(cekA)):
                                    if cekA[lupA] not in cekB:
                                        samadengan = pathB.split('\\')
                                        ket = 'TIDAK ADA DI "'+str(samadengan[-1])+'" PADA SHEET "'+str(self.m_sheetnameB.GetValue())+'"'
                                        hasil.append(ket)
                                    else :
                                        hasil.append(' ')
                                
                                dfA['HASIL PENGECEKAN SISTEM'] = hasil                                
                                
                                
                                self.df_result = pd.DataFrame()
                                self.nama_kolom = list(dfA)      
                                for kol in range(len(self.nama_kolom)):    
                                    self.m_datahasil.AppendTextColumn(self.nama_kolom[kol])
                                    
                                self.save = []
                                for lupA in range(len(cekA)):
                                    lis_temp = list(dfA.iloc[lupA])
                                    found = []
                                    self.dup_save = []
                                    for x in range(len(lis_temp)):
                                        temp = "'"+str(lis_temp[x])
                                        found.append(str(lis_temp[x]))
                                        self.dup_save.append(temp)
                                        
                                    self.save.append(self.dup_save)
                                    self.m_datahasil.AppendItem(found)
                                    found.clear()                                    
                                
                                
                                self.btn_unduh.Enable()                                
                                
                            
            except FileNotFoundError :
                msg = wx.MessageDialog(self, 'Anda belum upload file', 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy() 
                
            
            except ValueError as err:
                pesan = 'Terjadi kesalahan!!\n\nError : '+str(err)+'\n\nPeriksa dan sesuaikan kembali inputan anda'
                msg = wx.MessageDialog(self, pesan, 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()
                
                
            except:
                msg = wx.MessageDialog(self, 'Mohon maaf, Terjadi kesalahan\nJalankan ulang', 'Peringatan', wx.OK_DEFAULT)
                ans = msg.ShowModal()
                msg.Destroy()
                          
            
            
    
    
    
    def OnReset(self, e):
        self.btn_unduh.Disable()
        self.m_datahasil.DeleteAllItems()
        self.m_datahasil.ClearColumns()
        self.btn_reset.Disable()
        self.btn_tdkada.Disable()
        self.btn_periksa.Disable()
        self.daftarfeature.m_scrolledWindow8.Destroy()
        self.daftarfeature.indi = 1
        self.m_headersetB.SetValue(1)
        self.m_sheetnameB.SetLabel("")
        self.m_checkBox1.SetValue(False)
        self.m_uploadfileB.SetPath("")
        self.m_headersetA.SetValue(1)
        self.m_sheetnameA.SetLabel("")
        self.daftarfeatureA.m_scrolledWindow8.Destroy()
        self.daftarfeatureA.indi = 1        
        self.m_uploadfileA.SetPath("")
        
    def OnCheck (self, e):
        cb = e.GetEventObject()
        if self.m_uploadfileA.GetPath() == "" :
            msg = wx.MessageDialog(self, 'Anda belum upload file', 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy
            cb.SetValue(False)
        else :
            if cb.GetValue() == True :
                self.m_uploadfileB.SetPath(self.m_uploadfileA.GetPath())
            else :
                self.m_uploadfileB.SetPath("")
            
    
    def OnSubmit(self, e): 
        luping.clear()
        lupingA.clear()
        pathA = self.m_uploadfileA.GetPath()
        pathB = self.m_uploadfileB.GetPath()
        self.btn_periksa.Disable()
        self.btn_reset.Disable()
        self.btn_unduh.Disable()
        self.btn_tdkada.Disable()
        self.m_datahasil.DeleteAllItems()
        self.m_datahasil.ClearColumns()        
        try:
            if self.m_sheetnameA.IsEmpty():
                dfA = pd.read_excel(pathA, engine='openpyxl', header=int(self.m_headersetA.GetValue()-1))
                if self.m_sheetnameB.IsEmpty():
                    dfB = pd.read_excel(pathB, engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                    kol = list(dfB)
                    kolA = list(dfA)
                    
                    for y in range(len(kolA)):
                        lupingA.append(kolA[y])                
                    
                    for x in range(len(kol)):    
                        luping.append(kol[x])
                else :
                    sheetB = self.m_sheetnameB.GetValue()
                    sheetB = sheetB.split(", ")                
                    if len(sheetB) > 1 :
                        dfB = pd.read_excel(pathB, sheet_name=str(sheetB[0]), engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                        kol = list(dfB)
                        kolA = list(dfA)
                        
                        for y in range(len(kolA)):
                            lupingA.append(kolA[y])                    
                        
                        for x in range(len(kol)):    
                            luping.append(kol[x])
                            
                    else :
                        dfB = pd.read_excel(pathB, sheet_name=str(self.m_sheetnameB.GetValue()), engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                        kol = list(dfB)
                        kolA = list(dfA)
                        
                        for y in range(len(kolA)):
                            lupingA.append(kolA[y])                    
                        
                        for x in range(len(kol)):    
                            luping.append(kol[x])                
                
            else :
                sheetA = self.m_sheetnameA.GetValue()
                sheetA = sheetA.split(", ")
                if len(sheetA) > 1 :
                    msg = wx.MessageDialog(self, 'Sheet file satu hanya diisi satu sheet', 'Peringatan', wx.OK_DEFAULT)
                    ans = msg.ShowModal()
                    msg.Destroy()
                else:
                    dfA = pd.read_excel(pathA, sheet_name=str(self.m_sheetnameA.GetValue()), engine='openpyxl', header=int(self.m_headersetA.GetValue()-1))
                    if self.m_sheetnameB.IsEmpty():
                        dfB = pd.read_excel(pathB, engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                        kol = list(dfB)
                        kolA = list(dfA)
                        
                        for y in range(len(kolA)):
                            lupingA.append(kolA[y])                
                        
                        for x in range(len(kol)):    
                            luping.append(kol[x])
                            
                    else :
                        sheetB = self.m_sheetnameB.GetValue()
                        sheetB = sheetB.split(", ")                
                        if len(sheetB) > 1 :
                            dfB = pd.read_excel(pathB, sheet_name=str(sheetB[0]), engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                            kol = list(dfB)
                            kolA = list(dfA)
                            
                            for y in range(len(kolA)):
                                lupingA.append(kolA[y])                    
                            
                            for x in range(len(kol)):    
                                luping.append(kol[x])
                                
                        else :
                            dfB = pd.read_excel(pathB, sheet_name=str(self.m_sheetnameB.GetValue()), engine='openpyxl', header=int(self.m_headersetB.GetValue()-1))
                            kol = list(dfB)
                            kolA = list(dfA)
                            
                            for y in range(len(kolA)):
                                lupingA.append(kolA[y])                    
                            
                            for x in range(len(kol)):    
                                luping.append(kol[x])                    
                        
                
            
                    
                                
                        
                                
            self.btn_periksa.Enable()
            self.btn_tdkada.Enable()
            self.btn_reset.Enable()
            self.daftarfeatureA.ulang()
            self.daftarfeature.ulang()
            #self.luping.clear()
            
        except FileNotFoundError :
            msg = wx.MessageDialog(self, 'Anda belum upload file', 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy() 
            self.btn_periksa.Disable()
            self.btn_reset.Disable()
        
        except ValueError as err:
            pesan = 'Terjadi kesalahan!!\n\nError : '+str(err)+'\n\nPeriksa dan sesuaikan kembali inputan anda'
            msg = wx.MessageDialog(self, pesan, 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            self.btn_periksa.Disable()
            self.btn_reset.Disable()
            
        except:
            msg = wx.MessageDialog(self, 'Mohon maaf, Terjadi kesalahan\nJalankan ulang', 'Peringatan', wx.OK_DEFAULT)
            ans = msg.ShowModal()
            msg.Destroy()
            self.btn_periksa.Disable()
            self.btn_tdkada.Enable()
            self.btn_reset.Disable()       
    def onSave(self, e):
        
        m_save = pd.DataFrame(self.save, columns=self.nama_kolom)
        
        return m_save.to_excel(e) 
    
    def OnUnduh(self, event):
    
        with wx.FileDialog(self, "Save Excel file", wildcard="Excel files (*.xlsx)|*.xlsx",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
    
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                fileDialog.Close()
                
            else:
                pathname = fileDialog.GetPath()
                try:            
                    with open(pathname, 'wb') as file:
                        self.onSave(file)
                    msg = wx.MessageDialog(self, 'Berhasil disimpan', 'Information', wx.OK_DEFAULT)
                    ans = msg.ShowModal()
                    msg.Destroy()                    
                except IOError:
                    wx.LogError("Tidak dapat menyimpan data saat ini dalam file '%s'." % pathname)    



    
    
    
class MyFrame ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = "APLIKASI CEK DUPLIKASI", pos = wx.DefaultPosition, size = wx.Size( 1200,750 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
        self.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_BTNHIGHLIGHT ) )
        self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_BTNFACE ) )

        bSizer6 = wx.BoxSizer( wx.VERTICAL )

        fgSizer5 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer5.SetFlexibleDirection( wx.BOTH )
        fgSizer5.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        bSizer3 = wx.BoxSizer( wx.VERTICAL )

        self.m_buttonduplikat = wx.Button( self, wx.ID_ANY, u"Cek Duplikasi", wx.DefaultPosition, wx.Size( 100,30 ), 0 )
        bSizer3.Add( self.m_buttonduplikat, 0, wx.ALL, 5 )
        self.m_buttonduplikat.Bind(wx.EVT_BUTTON, self.onDuplikat)

        self.m_buttoncompare = wx.Button( self, wx.ID_ANY, u"Compare", wx.DefaultPosition, wx.Size( 100,30 ), 0 )
        bSizer3.Add( self.m_buttoncompare, 0, wx.ALL, 5 )
        self.m_buttoncompare.Bind(wx.EVT_BUTTON, self.onCompare)


        fgSizer5.Add( bSizer3, 1, wx.ALL|wx.EXPAND, 5 )

        self.m_paneltemp = wx.Panel( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
        b_paneltemp = wx.BoxSizer( wx.HORIZONTAL )

        self.m_duplikat = PanelDuplikasi(self.m_paneltemp)
        b_paneltemp.Add(self.m_duplikat,1, wx.EXPAND|wx.ALL, 5)
        
        self.m_compare = PanelCompare(self.m_paneltemp)
        b_paneltemp.Add(self.m_compare,0, wx.EXPAND|wx.ALL, 5)
        self.m_compare.Hide()
        

        self.m_paneltemp.SetSizer( b_paneltemp )
        self.m_paneltemp.Layout()
        b_paneltemp.Fit( self.m_paneltemp )
        fgSizer5.Add( self.m_paneltemp, 1, wx.EXPAND |wx.ALL, 5 )


        bSizer6.Add( fgSizer5, 1, wx.ALL|wx.EXPAND, 5 )


        self.SetSizer( bSizer6 )
        self.Layout()

        self.Centre( wx.BOTH )

    def onCompare(self, e):
        self.m_duplikat.Hide()
        self.m_compare.Show()

    def onDuplikat(self, e):
        self.m_compare.Hide()
        self.m_duplikat.Show()

    def __del__( self ):
        pass


app = wx.App()  
frame = MyFrame(None)
app.SetTopWindow(frame)
frame.Show()
app.MainLoop()    