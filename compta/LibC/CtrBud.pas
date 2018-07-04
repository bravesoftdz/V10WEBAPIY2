{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CtrBud;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, Menus, StdCtrls, ExtCtrls, Hcompte,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} 
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  Hctrls, ComCtrls, Grids, Buttons, Ent1, Hent1, UObjFiltres, {SG6 04/01/05 Gestion Filtres V6 FQ 15145}
  HStatus, HTB97, HPanel, UiUtil ;

Procedure ControleBudget(UnFb : TFichierBase) ;

Type TInfoCtrBud = Class
         Libelle : String ;
         Compte  : String ;
         Exclu   : String ;
       End ;

type
  TFCtrBud = class(TForm)
    FListe: THGrid;
    Pages: TPageControl;
    ChoixEtat: TTabSheet;
    Bevel1: TBevel;
    TCpte: TLabel;
    Tcpt1: TLabel;
    TCBud1: THLabel;
    TCBud2: THLabel;
    C1: THCpteEdit;
    C2: THCpteEdit;
    CbTous: TCheckBox;
    RgChoix: TRadioGroup;
    PopZ: TPopupMenu;
    HMTrad: THSystemMenu;
    FindDialog: TFindDialog;
    HM: THMsgBox;
    BBud: THBitBtn;
    BGen: THBitBtn;
    TCbAxe: TLabel;
    CbAxe: THValComboBox;
    PFiltres: TToolWindow97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    BFiltre: TToolbarButton97;
    CBud1: THCpteEdit;
    CBud2: THCpteEdit;
    POPF: TPopupMenu;
    BImprimer: TToolbarButton97;
    BOuvrir: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BStop: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BReduire: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    Dock971: TDock97;
    procedure BAnnulerClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CbAxeChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RgChoixClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BChercheClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure BBudClick(Sender: TObject);
    procedure BGenClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    //SG6 04/01/05 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    UnFb,fb1,fb2 : TFichierBase ;
    FNomFiltre : String ;
    WMinX,WMinY : Integer ;
    FirstFind,Stop : Boolean ;
    ListeBud,ListeGen : TStringList ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure PositionneLabel ;
    Procedure PositionneRequete(var StBud,StGen : String);
    Procedure VideLaListe ;
    Procedure BloqueControle(InRun : Boolean) ;
    Function  CreerUnObj(Q : TQuery) : TInfoCtrBud ;
//    Procedure RempliLesListes(Q : TQuery) ;
    Procedure ControleBudget ;
    Procedure ControleCompta ;
    Function  FaitRequeteCpta(Compte,CompteEx : String ; LeFb : TFichierBase ; SurTableLibre : Boolean) : String ;
    Procedure TraiteRequeteCpta(Ind : Integer ; LeFb : TFichierBase; Sql : String);
    Function  TestStop : Boolean ;
    Function  ChercheCompteDansBud(CptBud,ComptGene : String) : Boolean ;
    Function  ChercheCompte(Lig : Integer) : String ;
    function HasNotGeneral(Compte : String) : Boolean;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
    CPProcMetier,
    CPProcGen,
  {$ENDIF MODENT1}
     CalCole, //
     BUDGENE_TOM, // FicheBudgene
     BUDSECT_TOM, // FicheBudSect
{$IFDEF EAGLCLIENT}
     UtileAGl,    // PrintDBGrid
{$ELSE}
     PrintDBG, // PrintDBGrid
{$ENDIF}
     CPGeneraux_TOM,
     CPSection_TOM, HDB ;

Procedure ControleBudget(UnFb : TFichierBase) ;
var FCtrBud : TFCtrBud ;
    PP : THPanel ;
BEGIN
FCtrBud:=TFCtrBud.Create(Application) ;
FCtrBud.FNomFiltre:='CTRLBUD'+IntToStr(Ord(UnFb)) ; FCtrBud.UnFb:=UnFb ;
if UnFb=fbBudGen then FCtrBud.HelpContext:=15126000 Else FCtrBud.HelpContext:=15146000 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FCtrBud.ShowModal ;
    Finally
     FCtrBud.Free ;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FCtrBud,PP) ;
   FCtrBud.Show ;
   END ;
END ;

procedure TFCtrBud.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCtrBud.BAnnulerClick(Sender: TObject);
begin
  //SG6 04/01/05 Vide le panel;
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCtrBud.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, '');

WMinX:=Width ; WMinY:=Height ;
ListeBud:=TStringList.Create ; ListeGen:=TStringList.Create ;
end;

procedure TFCtrBud.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG5 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Free;

FListe.VidePile(False) ; VideLaListe ; ListeBud.Free ; ListeGen.Free ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCtrBud.FormShow(Sender: TObject);
begin
//SG6 04/01/2005 Gestion Filtres V6 FQ 15145
ObjFiltre.FFI_TABLE :=FNomFiltre;
ObjFiltre.Charger;

PositionneLabel ; RgChoix.ItemIndex:=0 ;
Case Unfb of
    fbBudgen : BEGIN BBud.Hint:=HM.Mess[16] ; BGen.Hint:=HM.Mess[18] ; END ;
    else BEGIN BBud.Hint:=HM.Mess[17] ; BGen.Hint:=HM.Mess[19] ; END ;
  End ;
end;

Procedure TFCtrBud.PositionneLabel ;
Var St1,St2 : String ;
BEGIN
RgChoix.Items.Clear ;
if UnFb=fbBudGen then
   BEGIN
   Caption:=HM.Mess[4] ;
   FListe.Cells[FListe.ColCount-1,0]:=HM.Mess[13] ;
   RgChoix.Items.Add(HM.Mess[0]) ; RgChoix.Items.Add(HM.Mess[1]) ;
   TCBud1.Caption:=HM.Mess[6] ; TCpte.Caption:=HM.Mess[7] ;
   CBud1.ZoomTable:=tzBudgen ; CBud2.ZoomTable:=tzBudgen ;
   C1.ZoomTable:=tzGeneral ; C2.ZoomTable:=tzGeneral ;
   CbAxe.Enabled:=False ; TCbAxe.Enabled:=False ;
   CbAxe.Visible:=False ; TCbAxe.Visible:=False ;
   fb1:=fbBudGen ; fb2:=fbGene ;
   PremierDernier(fb1,St1,St2) ; CBud1.Text:=St1 ; CBud2.Text:=St2 ;
   PremierDernier(fb2,St1,St2) ; C1.Text:=St1 ; C2.Text:=St2 ;
   END else
   BEGIN
   Caption:=HM.Mess[5] ;
   FListe.Cells[FListe.ColCount-1,0]:=HM.Mess[14] ;
   RgChoix.Items.Add(HM.Mess[2]) ; RgChoix.Items.Add(HM.Mess[3]) ;
   TCBud1.Caption:=HM.Mess[8] ; TCpte.Caption:=HM.Mess[9] ;
   CbTous.Caption:=HM.Mess[20] ;
   CbAxe.Enabled:=True ; TCbAxe.Enabled:=True ;
   If CbAxe.Values.Count>0 Then CbAxe.Value:=CbAxe.Values[0] ;
   END ;
UpdateCaption(Self) ;
END ;

procedure TFCtrBud.CbAxeChange(Sender: TObject);
Var St1,St2 : String ;
begin
Case CbAxe.Value[2] of
     '1' : BEGIN
           CBud1.ZoomTable:=tzBudSec1 ; CBud2.ZoomTable:=tzBudSec1 ;
           C1.ZoomTable:=tzSection ; C2.ZoomTable:=tzSection ;
           fb1:=fbBudSec1 ; fb2:=fbAxe1 ;
           END ;
     '2' : BEGIN
           CBud1.ZoomTable:=tzBudSec2 ; CBud2.ZoomTable:=tzBudSec2 ;
           C1.ZoomTable:=tzSection2 ; C2.ZoomTable:=tzSection2 ;
           fb1:=fbBudSec2 ; fb2:=fbAxe2 ;
           END ;
     '3' : BEGIN
           CBud1.ZoomTable:=tzBudSec3 ; CBud2.ZoomTable:=tzBudSec3 ;
           C1.ZoomTable:=tzSection3 ; C2.ZoomTable:=tzSection3 ;
           fb1:=fbBudSec3 ; fb2:=fbAxe3 ;
           END ;
     '4' : BEGIN
           CBud1.ZoomTable:=tzBudSec4 ; CBud2.ZoomTable:=tzBudSec4 ;
           C1.ZoomTable:=tzSection4 ; C2.ZoomTable:=tzSection4 ;
           fb1:=fbBudSec4 ; fb2:=fbAxe4 ;
           END ;
     '5' : BEGIN
           CBud1.ZoomTable:=tzBudSec5 ; CBud2.ZoomTable:=tzBudSec5 ;
           C1.ZoomTable:=tzSection5 ; C2.ZoomTable:=tzSection5 ;
           fb1:=fbBudSec5 ; fb2:=fbAxe5 ;
           END ;
  End ;
PremierDernier(fb1,St1,St2) ; CBud1.Text:=St1 ; CBud2.Text:=St2 ;
PremierDernier(fb2,St1,St2) ; C1.Text:=St1 ; C2.Text:=St2 ;
end;

procedure TFCtrBud.RgChoixClick(Sender: TObject);
Var St1,St2 : String ;
begin
PremierDernier(fb1,St1,St2) ; CBud1.Text:=St1 ; CBud2.Text:=St2 ;
PremierDernier(fb2,St1,St2) ; C1.Text:=St1 ; C2.Text:=St2 ;
Case RgChoix.ItemIndex of
     0 : BEGIN
         C1.Enabled:=False ; C2.Enabled:=False ; TCpte.Enabled:=False ; Tcpt1.Enabled:=False ;
         CBud1.Enabled:=True ; CBud2.Enabled:=True ; TCBud1.Enabled:=True ; TCBud2.Enabled:=True ;
         CbTous.Enabled:=False ;
         END ;
     1 : BEGIN
         C1.Enabled:=True ; C2.Enabled:=True ; TCpte.Enabled:=True ; Tcpt1.Enabled:=True ;
         CBud1.Enabled:=False ; CBud2.Enabled:=False ; TCBud1.Enabled:=False ; TCBud2.Enabled:=False ;
         CbTous.Enabled:=True ;
         END ;
   End ;
end;

procedure TFCtrBud.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FirstFind); end ;

procedure TFCtrBud.BRechercherClick(Sender: TObject);
begin FirstFind:=true; FindDialog.Execute ; end;

Procedure TFCtrBud.PositionneRequete(var StBud,StGen : String);
BEGIN
Case RgChoix.ItemIndex of
     0 : BEGIN
         if UnFb=fbBudGen then
            BEGIN
            StBud:='Select BG_BUDGENE,BG_LIBELLE,BG_COMPTERUB,BG_EXCLURUB From BUDGENE Where BG_BUDGENE>="'+CBud1.Text+'" And '+
                   'BG_BUDGENE<="'+CBud2.Text+'"' ;
            StGen:='Select G_GENERAL,G_LIBELLE From GENERAUX' ;
            END else
            BEGIN
            StBud:='Select BS_BUDSECT,BS_LIBELLE,BS_SECTIONRUB,BS_EXCLURUB From BUDSECT '+
                   'Where BS_BUDSECT>="'+CBud1.Text+'" And BS_BUDSECT<="'+CBud2.Text+'" And BS_AXE="'+CbAxe.Value+'"' ;
            StGen:='Select S_SECTION,S_LIBELLE From SECTION Where S_AXE="'+CbAxe.Value+'"' ;
            END ;
         END ;
     1 : BEGIN
         if UnFb=fbBudGen then
            BEGIN
            StBud:='Select BG_BUDGENE,BG_LIBELLE,BG_COMPTERUB,BG_EXCLURUB From BUDGENE' ;
            StGen:='Select G_GENERAL,G_LIBELLE From GENERAUX Where '+
                   'G_GENERAL>="'+C1.Text+'" And G_GENERAL<="'+C2.Text+'"' ;
            END else
            BEGIN
            StBud:='Select BS_BUDSECT,BS_LIBELLE,BS_SECTIONRUB,BS_EXCLURUB From BUDSECT '+
                   'Where BS_AXE="'+CbAxe.Value+'"' ;
            StGen:='Select S_SECTION,S_LIBELLE From SECTION Where S_AXE="'+CbAxe.Value+'" '+
                   'And S_SECTION>="'+C1.Text+'" And S_SECTION<="'+C2.Text+'"' ;
            END ;
         END ;
    End ;
END ;

Procedure TFCtrBud.VideLaListe ;
Var i : Integer ;
BEGIN
for i:=0 to ListeBud.Count-1 do if TObject(ListeBud.Objects[i])<>Nil then TObject(ListeBud.Objects[i]).Free ;
ListeBud.Clear ;
for i:=0 to ListeGen.Count-1 do if TObject(ListeGen.Objects[i])<>Nil then TObject(ListeGen.Objects[i]).Free ;
ListeGen.Clear ;
END ;

procedure TFCtrBud.BStopClick(Sender: TObject);
begin Stop:=True ; end;

Function TFCtrBud.TestStop : Boolean ;
BEGIN
Application.ProcessMessages ; Result:=False ;
if Not Stop then Exit ;
Stop:=False ;
if HM.Execute(15,Caption,'')=mrYes then Result:=True ;
END ;

Procedure TFCtrBud.BloqueControle(InRun : Boolean) ;
BEGIN
Pages.Enabled:=Not InRun ; PFiltres.Enabled:=Not InRun ; BAgrandir.Enabled:=Not InRun ;
BReduire.Enabled:=Not InRun ; BRechercher.Enabled:=Not InRun ;
BMenuZoom.Enabled:=Not InRun ; 
END ;

Function TFCtrBud.CreerUnObj(Q : TQuery) : TInfoCtrBud ;
Var X : TInfoCtrBud ;
BEGIN
Result:=Nil ;
if Q=Nil then Exit ;
X:=TInfoCtrBud.Create ;
X.Libelle:=Q.Fields[1].AsString ;
if Q.FieldCount>2 then
   BEGIN
   X.Compte:=Q.Fields[2].AsString ; X.Exclu:=Q.Fields[3].AsString ;
   END ;
Result:=X ;
END ;

{Procedure TFCtrBud.RempliLesListes(Q : TQuery) ;
BEGIN
if UpperCase(Q.Name)='QBUD' then
   While Not Q.Eof do BEGIN ListeBud.AddObject(Q.Fields[0].AsString,CreerUnObj(Q)) ; Q.Next ; END
else
   While Not Q.Eof do BEGIN ListeGen.AddObject(Q.Fields[0].AsString,CreerUnObj(Q)) ; Q.Next ; END ;
END ;}

procedure TFCtrBud.BChercheClick(Sender: TObject);
var
  QBud, QCpta : TQuery;
  StBud,StGen : String;
begin
  BloqueControle(True);
  PositionneRequete(StBud,StGen);
  FListe.VidePile(False);
  VideLaListe;

  QBud := OpenSQL(StBud, True);
  While Not QBud.Eof do begin
    ListeBud.AddObject(QBud.Fields[0].AsString,CreerUnObj(QBud));
    QBud.Next;
  end;
  Ferme(QBud);

  QCpta := OpenSQL(StGen, True);
  While Not QCpta.Eof do begin
    ListeGen.AddObject(QCpta.Fields[0].AsString, CreerUnObj(QCpta));
    QCpta.Next;
  end;

Stop:=False ;
Case RgChoix.ItemIndex of
     0: BEGIN
        InitMove(ListeBud.Count,'') ;
        ControleBudget ; FiniMove ;
        END ;
     1: BEGIN
        InitMove(ListeGen.Count,'') ;
        ControleCompta ; FiniMove ;
        END ;
  End ;
BloqueControle(False) ;
end;

Procedure TFCtrBud.ControleBudget ;
Var i : Integer ;
    Cpt1,CpEx1 : String ;
    Sql : String ;
BEGIN
for i:=0 to ListeBud.Count-1 do
    BEGIN
    MoveCur(False) ;
    FListe.Cells[0,FListe.RowCount-1]:=ListeBud.Strings[i] ;
    FListe.Cells[1,FListe.RowCount-1]:=TInfoCtrBud(ListeBud.Objects[i]).Libelle ;
    Cpt1 := TInfoCtrBud(ListeBud.Objects[i]).Compte;
    CpEx1:= TInfoCtrBud(ListeBud.Objects[i]).Exclu;

    Sql:=FaitRequeteCpta(Cpt1,CpEx1,fb2,False) ;
    if Sql<>'' then TraiteRequeteCpta(0,fb2, Sql);

    Sql:=FaitRequeteCpta(CpEx1,'',fb2,False) ;
    if Sql<>'' then TraiteRequeteCpta(1,fb2, Sql);

    if TestStop then Break ;
    END ;
FiniMove ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
END ;

Procedure TFCtrBud.ControleCompta ;
Var i,j : Integer ;
    PremierTour : Boolean ;
    Compte,CptBud,CptEx : String ;
BEGIN
for j:=0 to ListeGen.Count-1 do
  BEGIN
  MoveCur(False) ; PremierTour:=True ;
  for i:=0 to ListeBud.Count-1 do
    BEGIN
    if PremierTour then
       BEGIN
       Compte:=ListeGen.Strings[j] ;
       FListe.Cells[0,FListe.RowCount-1]:=Compte ;
       FListe.Cells[1,FListe.RowCount-1]:=TInfoCtrBud(ListeGen.Objects[j]).Libelle ;
       if UnFb in [fbAxe1..fbAxe5] then
          BEGIN
          if VH^.Cpta[UnFb].Chantier then Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[12]
                                    else Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[11] ;
          END else
          BEGIN
          if HasNotGeneral(Compte) then Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[11]
                                   else Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[12] ;
          END ;
       PremierTour:=False ;
       END ;
       CptBud:=TInfoCtrBud(ListeBud.Objects[i]).Compte ;
       CptEx:=TInfoCtrBud(ListeBud.Objects[i]).Exclu ;
       if ChercheCompteDansBud(CptBud,Compte)then
          BEGIN
          FListe.Cells[2,FListe.RowCount-1]:=ListeBud.Strings[i] ;
          FListe.Cells[3,FListe.RowCount-1]:=TInfoCtrBud(ListeBud.Objects[i]).Libelle ;
          FListe.Cells[4,FListe.RowCount-1]:=HM.Mess[11] ;
          FListe.RowCount:=FListe.RowCount+1 ;
          END ;
       if (CptEx<>'') And (Pos(';',CptEx)>1) then
          BEGIN
          if ChercheCompteDansBud(CptEx,Compte)then
             BEGIN
             FListe.Cells[2,FListe.RowCount-1]:=ListeBud.Strings[i] ;
             FListe.Cells[3,FListe.RowCount-1]:=TInfoCtrBud(ListeBud.Objects[i]).Libelle ;
             FListe.Cells[4,FListe.RowCount-1]:=HM.Mess[12] ;
             FListe.RowCount:=FListe.RowCount+1 ;
             END ;
          END ;
       END ;
    if CbTous.Checked then
       if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1 ;
    if TestStop then Break ;
    END ;
if Fliste.RowCount>2 then Fliste.RowCount:=Fliste.RowCount-1 ;
if(Fliste.RowCount=2) And (Fliste.Cells[2,1]='') then
   for i:=0 to Fliste.ColCount-1 do Fliste.Cells[i,1]:='' ;
END ;

Function TFCtrBud.FaitRequeteCpta(Compte,CompteEx : String ; LeFb : TFichierBase ; SurTableLibre : Boolean) : String ;
Var Sql,Cod,Lib,Table,Where,WhereAnd : String ;
BEGIN
Result:='' ; Sql:='' ;
Case LeFb of
    fbGene : BEGIN Cod:='G_GENERAL' ; Lib:='G_LIBELLE' ; Table:='GENERAUX' ; WhereAnd:='' ; END ;
    fbAxe1..fbAxe2 : BEGIN Cod:='S_SECTION' ; Lib:='S_LIBELLE' ; Table:='SECTION' ; WhereAnd:=' And S_AXE="'+CbAxe.Value+'"' ; END ;
    End ;
Where:=AnalyseCompte(Compte,LeFb,False,SurTableLibre, True, False); // 12102
if Where<>'' then
   BEGIN
   Sql:='Select '+Cod+','+Lib+' From '+Table+'' ;
   Sql:=Sql+' Where '+Where ;
   END else Exit ;
if Compteex<>'' then
   BEGIN
   Where:=AnalyseCompte(Compteex,Lefb,True,FALSE, True, False); // 12102
   if Where<>'' then Sql:=Sql+' And '+Where ;
   END ;
Sql:=Sql+WhereAnd+' Order by '+Cod ;
Result:=Sql ;
END ;

Procedure TFCtrBud.TraiteRequeteCpta(Ind : Integer ; LeFb : TFichierBase; Sql : String);
var
  QCpta : TQuery;
BEGIN
  QCpta := OpenSQL(Sql, True);

  while Not QCpta.Eof do begin
    FListe.Cells[2,FListe.RowCount-1]:=QCpta.Fields[0].AsString ;
    FListe.Cells[3,FListe.RowCount-1]:=QCpta.Fields[1].AsString ;
    FListe.Cells[4,FListe.RowCount-1]:=HM.Mess[11+Ind] ;
    if Not (Lefb in [fbAxe1..fbAxe2]) then begin
      if HasNotGeneral(QCpta.Fields[0].AsString) then Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[11]
                                                 else Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[12] ;
      end
    else begin
      if VH^.Cpta[LeFb].Chantier then Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[12]
                                 else Fliste.Cells[5,FListe.RowCount-1]:=HM.Mess[11] ;
    end;
    QCpta.Next;
    FListe.RowCount:=FListe.RowCount+1;
  end;
  Ferme(QCpta);
end;

Function TFCtrBud.ChercheCompteDansBud(CptBud,ComptGene : String) : Boolean ;
Var St,StTemp : String ;
BEGIN
Result:=False ;
While CptBud<>'' do
  BEGIN
  St:=ReadTokenSt(CptBud) ; StTemp:='' ;
  if St='' then Continue ;
  if Pos(':',St)>0 then BEGIN StTemp:=Copy(St,Pos(':',St)+1,200) ; Delete(St,Pos(':',St),200) ; END ;
  if StTemp<>'' then
     BEGIN
     St:=BourrelaDonc(St,fb2) ; StTemp:=BourreLaDonc(StTemp,fb2) ;
     if(ComptGene>=St) And (ComptGene<=StTemp) then
        BEGIN Result:=True ; Exit ; END ;
     END else
     BEGIN
     if Length(St)<=VH^.Cpta[fb2].Lg then
        if(Pos(St,ComptGene)=1)Or(St=ComptGene) then
           BEGIN Result:=True ; Exit ; END ;
     END ;
  END ;
END ;

Function TFCtrBud.ChercheCompte(Lig : Integer) : String ;
Var i : Integer ;
BEGIN
Result:='' ;
for i:=Lig DownTo 1 do if FListe.Cells[0,i]<>'' then BEGIN Result:=FListe.Cells[0,i] ; Exit ; END ;
END ;

procedure TFCtrBud.BBudClick(Sender: TObject);
Var Cpt : String ;
begin
if FListe.Cells[0,1]='' then Exit ;
Case RgChoix.ItemIndex of
     0 : BEGIN
         if FListe.Cells[0,FListe.Row]='' then Cpt:=ChercheCompte(FListe.Row)
                                          else Cpt:=FListe.Cells[0,FListe.Row] ;
         if UnFb=fbBudGen then FicheBudgene(Nil,'',Cpt,taConsult,0)
                          else FicheBudsect(Nil,'',Cpt,taConsult,0) ;
         END ;
     1 : BEGIN
         if UnFb=fbBudGen then FicheBudgene(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0)
                          else FicheBudsect(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0) ;
         END ;
   End ;
end;

procedure TFCtrBud.BGenClick(Sender: TObject);
Var Cpt : String ;
begin
if FListe.Cells[0,1]='' then Exit ;
Case RgChoix.ItemIndex of
     0 : BEGIN
         if UnFb=fbBudGen then FicheGene(Nil,'',FListe.Cells[2,FListe.Row],taConsult,0)
                          else FicheSection(Nil,CbAxe.Value,FListe.Cells[2,FListe.Row],taConsult,0) ;
         END ;
     1 : BEGIN
         if FListe.Cells[0,FListe.Row]='' then Cpt:=ChercheCompte(FListe.Row)
                                          else Cpt:=FListe.Cells[0,FListe.Row] ;
         if UnFb=fbBudGen then FicheGene(Nil,'',Cpt,taConsult,0)
                          else FicheSection(Nil,CbAxe.Value,Cpt,taConsult,0) ;
         END ;
   End ;
end;

procedure TFCtrBud.FListeDblClick(Sender: TObject);
begin
Case RgChoix.ItemIndex of
     0 :BBudClick(Nil) ;
     1 :BGenClick(Nil) ;
   End ;
end;

procedure TFCtrBud.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFCtrBud.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFCtrBud.BImprimerClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  T, F : Tob;
  i : Integer;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  SourisSablier;
  T := TOB.Create('non', nil, -1);
  for i := 1 to FListe.RowCount-1 do begin
    F := TOB.Create('non', T, -1);
    F.AddChampSup('C1', False);
    F.AddChampSup('L1', False);
    F.AddChampSup('C2', False);
    F.AddChampSup('L2', False);
    F.AddChampSup('EXCLU', False);
    F.AddChampSup('MOUV', False);
    if i=1 then begin
      F.AddChampSup('TITRE', False);
      F.SetString('TITRE', Caption);
    end;
    F.SetString('C1', FListe.Cells[0, i]);
    F.SetString('L1', FListe.Cells[1, i]);
    F.SetString('C2', FListe.Cells[2, i]);
    F.SetString('L2', FListe.Cells[3, i]);
    F.SetString('EXCLU', FListe.Cells[4, i]);
    F.SetString('MOUV', FListe.Cells[5, i]);
  end;
  SourisNormale;
  LanceEtatTob('E','CST','CBU',T, True, False, False, nil, '', Caption, False, 0, '', 0, '');
  T.Free;
{$ELSE}
  PrintDBGrid(FListe,Nil, Caption,'');
{$ENDIF}
end;

procedure TFCtrBud.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFCtrBud.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ; 
end;

procedure TFCtrBud.BOuvrirClick(Sender: TObject);
begin
BBudClick(Nil) ;
end;

procedure TFCtrBud.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //SG6 04/01/05 Gestion des Filtres V6 FQ 15145
  if Key = VK_F9 then BChercheClick(nil);
end;

function TFCtrBud.HasNotGeneral(Compte: String): Boolean;
begin
  Result := not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+Compte+'" '+
                          'AND ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+Compte+'")) '+
                          'OR (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL="'+Compte+'")))');
end;

end.
