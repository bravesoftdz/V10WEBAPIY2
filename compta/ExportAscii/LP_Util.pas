unit LP_Util;

interface

uses ImgList, Controls, ComCtrls, StdCtrls, {Mask, }Hctrls, HTB97,  //XMG 05/04/04 ne sert à rien....
     Classes, HSysMenu, Forms, Mask ; //, Mask ; //XMG 15/04/04

Procedure ChargeChampsTree ( Tables : String ; var TreeCh : TTreeView )  ;

type
  TFLPChamps = class(TForm)
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BCancel: TToolbarButton97;
    BAide: TToolbarButton97;
    Images: TImageList;
    TChamp: THLabel;
    Champ: THCritMaskEdit;
    TreeCh: TTreeView;
    BChamp: TToolbarButton97;
    BLibelle: TToolbarButton97;
    procedure CMDialogKey(var Message: TCMDialogKey); message CM_DIALOGKEY ;
    procedure BLibelleClick(Sender: TObject);
    procedure ChampChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure TreeChClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    Procedure ChercheChamp ;
    Procedure RefreshTree ;
    Procedure SelectChamp ;
  public
    TreeO : TTreeView ;
    Tipe    : String ;
  end;


implementation

{$R *.DFM}

uses Hent1, MC_Erreur, Sysutils, Windows, Ed_Tools,
{$IFDEF eAGLClient}
     UTOB
{$ELSE}
     dbTables
{$ENDIF}
 ;


type
     TDecla = Class
       Champ,
       Libelle,
       TypeCh : String ;
       End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPChamps.RefreshTree ;
Var Node       : TTreeNode ;
Begin
TreeCh.Items.beginUpdate ;
Node:=TreeCH.items.GetFirstNode ;
while node<>Nil do
  Begin
  if BChamp.down then Node.Text:=TDecla(Node.Data).Champ else Node.Text:=TDecla(Node.Data).Libelle ;
  Node:=Node.GetNext ;
  End ;
TreeCh.Items.EndUpdate ;
ChercheChamp ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.BLibelleClick(Sender: TObject);
begin
RefreshTree ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPChamps.ChercheChamp ;
Var i : Integer ;
    Node : TTreeNode ;
    St   : String ;
Begin
if Champ.Text='' then exit ;
St:=Champ.Text ;
If (Length(St)>1) and (St[1]='[') then delete(St,1,1) ;
If (St[Length(St)]=']') then delete(St,Length(St),1) ;
Node:=TreeCh.Items.GetFirstNode ;
i:=-1 ;
While (Node<>nil) and (i=-1) do
  Begin
  if St=Copy(TDecla(Node.Data).Champ,1,length(St)) then i:=Node.absoluteindex ;
  Node:=Node.GetNext ;
  End ;
if i>=0 then TreeCh.Selected:=TreeCh.items[i] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.ChampChange(Sender: TObject);
begin
if Champ.Text<>'' then ChercheChamp
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.BValiderClick(Sender: TObject);
begin
SelectChamp ;
if Champ.Text='' then exit ;
ModalResult:=MrOk ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPChamps.SelectChamp ;
Var fils : TTreeNode ;
Begin
Fils:=TreeCh.Selected ;
if Fils.count>0 then Begin Champ.Text:='' ; exit ; End ;
Champ.Text:=TDecla(Fils.Data).Champ ;
Tipe:=TDecla(Fils.Data).TypeCh ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.TreeChClick(Sender: TObject);
begin
SelectChamp ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure CreerArbre ( var Pere : TTreeNode ; Table,Nom,Libelle,Tipe : String ; TreeCh : TTreeView ) ;
      ///////////////////////////////////////////////////////
      Function CreeAnnexe ( Pere: TTreeNode ; Nom,Libelle,Tipe : String ; TreeCh : TTreeView ) : TTreeNode ;
      Var Decla   : TDecla ;
      Begin
      Decla:=TDecla.Create ;
      Decla.Champ:=Nom ;
      Decla.Libelle:=Libelle ;
      Decla.TypeCh:=tipe ;
      Result:=TreeCh.Items.addChildObject(Pere,Libelle,Decla) ;
      End ;
      ///////////////////////////////////////////////////////
Var Fils     : TTreeNode ;
    LblTable : String ;
Begin
if Pere=nil then
   Begin
   if Table='SYSTEME' then LblTable:=TraduireMemoire(MC_MsgErrDefaut(1041)) else
   if Table='LPRINTER' then LblTable:=TraduireMemoire(MC_MsgErrDefaut(1042))
      else LblTable:=TabletoLibelle(Table)  ;
   Pere:=CreeAnnexe(nil,Table,LblTable,'',TreeCh) ;
   Pere.ImageIndex:=0 ;
   Pere.SelectedIndex:=0 ;
   Pere.ImageIndex:=0 ;
   End ;
Fils:=CreeAnnexe(Pere,Nom,Libelle,Tipe,TreeCh) ;
Fils.imageindex:=1 ;
Fils.Selectedindex:=1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure ChargeVariablesSpeciaux ( Var Pere : TTreeNode ; Table : String ; TreeCh : TTreeView ) ;
Begin
CreerArbre(Pere,Table,'TEX_PARTIEL',TraduireMemoire(MC_MsgErrDefaut(1043)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_TOTAL'  ,TraduireMemoire(MC_MsgErrDefaut(1044)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_OUVRE'  ,TraduireMemoire(MC_MsgErrDefaut(1045)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_DEBSLIP',TraduireMemoire(MC_MsgErrDefaut(1063)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_FINSLIP',TraduireMemoire(MC_MsgErrDefaut(1064)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_DEBVAL' ,TraduireMemoire(MC_MsgErrDefaut(1065)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_FINVAL' ,TraduireMemoire(MC_MsgErrDefaut(1066)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_DEBCHQ' ,TraduireMemoire(MC_MsgErrDefaut(1068)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_FINCHQ' ,TraduireMemoire(MC_MsgErrDefaut(1069)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_IMPBMP' ,TraduireMemoire(MC_MsgErrDefaut(1078)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_DEBLNSP',TraduireMemoire(MC_MsgErrDefaut(1079)),'',TreeCH) ;
CreerArbre(Pere,Table,'TEX_FINLNSP',TraduireMemoire(MC_MsgErrDefaut(1080)),'',TreeCH) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure ChargeParamSoc (var Pere : TTreeNode ; Table : String ; TreeCH : TTreeView );
Var Q                   : TQuery ;
    Nom,Libelle,Stype,Tipe,St : String ;
Begin
Q:=OpenSQL('SELECT SOC_NOM, SOC_DESIGN FROM PARAMSOC WHERE SOC_NOM like "SO_%"',TRUE) ;
While Not Q.Eof do
  Begin
  Nom:=Q.FindField('SOC_NOM').asString ;
  St:=Q.FindField('SOC_DESIGN').asString ;
  SType:=readtokenSt(St) ; Tipe:='' ;
  Libelle:=ReadTokenSt(St) ; Libelle:=ReadTokenSt(St) ;
  if (stype='A') or (stype='C') or (stype='L') then tipe:='VARCHAR' else
  if (stype='I') or (stype='N') then tipe:='INTEGER' else
  if (stype='B') then tipe:='BOOLEAN' else
  if (stype='D') then tipe:='DATE' else
  if (stype='R') then tipe:='DATE' else
  if (stype='F') or (stype='T') or (stype='P') then tipe:='DOUBLE' ;
  if (trim(Nom)<>'') and (Tipe<>'') then CreerArbre(Pere,Table,Nom,Libelle,Tipe,TreeCh) ;
  Q.Next ;
  End ;
Ferme(Q);
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure ChargeVariablesSysteme(Var Pere : TTreeNode ; Table : String ; TreeCh : TTreeView ) ;
Begin
CreerArbre(Pere,Table,'SYS_COPYRIGHT',TraduireMemoire(MC_MsgErrDefaut(1046)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_DATE',TraduireMemoire(MC_MsgErrDefaut(1047)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_DATEENTREE',TraduireMemoire(MC_MsgErrDefaut(1048)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_DATEVERSION',TraduireMemoire(MC_MsgErrDefaut(1049)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_HALLEY',TraduireMemoire(MC_MsgErrDefaut(1050)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_HEURE',TraduireMemoire(MC_MsgErrDefaut(1051)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_NOMUTILISATEUR',TraduireMemoire(MC_MsgErrDefaut(1052)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_PROGVERSION',TraduireMemoire(MC_MsgErrDefaut(1053)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_TITRE',TraduireMemoire(MC_MsgErrDefaut(1054)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_UTILISATEUR',TraduireMemoire(MC_MsgErrDefaut(1055)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_NATURE',TraduireMemoire(MC_MsgErrDefaut(1056)),'',TreeCH) ;
CreerArbre(Pere,Table,'SYS_MODELE',TraduireMemoire(MC_MsgErrDefaut(1057)),'',TreeCH) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure ChargeChampsTree ( Tables : String ; var TreeCh : TTreeView )  ;
Var St,St2,Pref : String ;
    NumTable,j  : Integer ;
    Pere        : TTreeNode ;
Begin
TreeCh.Items.beginUpdate ;
TreeCh.items.clear ;
St:=Tables ;
while St<>'' do
  Begin
  Pere:=nil ;
  St2:=ReadTokenSt(St) ;
  if St2='PARAMSOC' then ChargeParamSoc(Pere,St2,TreeCH) else
  if St2='LPRINTER' then ChargeVariablesSpeciaux(Pere,St2,TreeCH) else
     Begin
     Pref:=TableToPrefixe(St2) ;
     if Pref='' then NumTable:=-1 else NumTable:=PrefixeToNum(Pref) ;
     if NumTable>=1 then
        for j:=low(V_PGI.DEChamps[NumTable]) to high(V_PGI.DEChamps[NumTable]) do
            if trim(V_PGI.DEChamps[NumTable,j].Nom)<>'' then CreerArbre(Pere,St2,V_PGI.DEChamps[NumTable,j].Nom,V_PGI.DEChamps[NumTable,j].Libelle,V_PGI.DEChamps[NumTable,j].tipe,TreeCH) ;
     End ;
  End ;
Pere:=nil ; 
ChargeVariablesSysteme(Pere,'SYSTEME',TreeCh) ;
TreeCh.Items.EndUpdate ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
if Key = VK_VALIDE  then  bValiderClick(nil) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.CMDialogKey(var Message: TCMDialogKey);
var keyState : Short ;
    ParentForm : TForm ;
Begin
If not (ActiveControl Is TButton) Then
   If Message.Charcode = VK_RETURN Then
      Begin
      if IsInside(Self) then
         Begin
         KeyState:=GetKeyState(VK_SHIFT) ;
         if KeyState>=0 then
            Begin
            parentForm:=TForm(GetParentForm(Self)) ;
            NextControl(ParentForm,TRUE) ;
            Message.Result:=1 ;
            End ;
         End else Message.Charcode := VK_TAB ;
      End ;
if Message.Result=0 then inherited;
End;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPChamps.FormDestroy(Sender: TObject);
begin
VideTree(TreeCH) ;
TreeCH.items.Clear ; 
end;

end.
