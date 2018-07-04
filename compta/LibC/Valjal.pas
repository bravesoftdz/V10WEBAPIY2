{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/03/2003
Modifié le ... : 22/09/2004
Description .. : Passage en eAGL
Suite ........ : 
Suite ........ : - BPY le 22/09/2004 - Fiche n° 14635 + optimization
Mots clefs ... : 
*****************************************************************}
unit ValJal;

interface

uses Windows,
     Messages,
     SysUtils,
     Classes,
{$IFDEF EAGLCLIENT}
     UTOb,        // TQuery
{$ELSE}
     DB,          // TQuery
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     Graphics,
     Controls,
     Forms,
     Dialogs,
     Ent1,
     Hent1,
     StdCtrls,
     Hctrls,
     Buttons,
     ExtCtrls,
     hmsgbox,
     Grids,
     HSysMenu,
     HPanel,
     UiUtil,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     HTB97;

Procedure ValidationJournal(Validation : Boolean ) ;

Type TInfoSup = Class
        St : String ;
      End ;

type
  TFValJal = class(TForm)
    PExo: TPanel;
    TExo: THLabel;
    CbExo: THValComboBox;
    CbVal: TCheckBox;
    BCherche: TToolbarButton97;
    Panel2: TPanel;
    PJal: TPanel;
    PPer: TPanel;
    MsgBox: THMsgBox;
    CBDebDat: TComboBox;
    CBFinDate: TComboBox;
    PerJal: TComboBox;
    PerJal1: TComboBox;
    HMTrad: THSystemMenu;
    BTag: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    Bdetag: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Nb1: TLabel;
    Tex1: TLabel;
    HPB: TToolWindow97;
    Dock: TDock97;
    Fliste: THGrid;
    Fliste1: THGrid;
    iCritGlyph: TImage;
    iCritGlyphModified: TImage;
    procedure CbExoChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure Fliste1Click(Sender: TObject);
    procedure FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FlisteDblClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CbValClick(Sender: TObject);
  private
    FCritModified: Boolean;                                          {FP 10/11/2005 FQ14941}
    Validation : Boolean ;
    Exo : TExoDate ;
    NbrMois : Word ;
    PerioJal : String ;
    NbPerio,TotalSelec : Integer ;
    WMinX,WMinY : Integer ;
    procedure SetCritModified(const Value: Boolean);                 {FP 10/11/2005 FQ14941}
    Procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure InitExo ;
    Procedure RempliCbExo ;
    Procedure RempliLesDates ;
    Function  FiltreJal : String ;
    Procedure ChercheJournaux ;
    Procedure ExoValidOuiNon ;
    Procedure CompteElemSelectionner ;
//    Procedure VideCombo(Sender : TComBoBox) ;
    Procedure AfficheResultat ;
    Procedure GeleLesBoutons ;
    Procedure TagDetag(Avec : Boolean) ;
    Function  OnExecute : Boolean ;
    function  MajObjetJalVal(StCod : String) : string;
    Procedure MajDesTables ;
    Function  OnValideAussiLexo : Boolean ;
    Function  MajObjetJalDeVal(StCod : String) : String ;

    property CritModified: Boolean read FCritModified write SetCritModified; {FP 10/11/2005 FQ14941}
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

uses
  ParamSoc,
{$IFDEF EAGLCLIENT}
// A FAIRE  Voir BImprimerClick
{$ELSE}
     PrintDBG,
{$ENDIF}
     Constantes,
     UtilPgi ;

{$R *.DFM}

Procedure ValidationJournal(Validation : Boolean ) ;
var FValJal: TFValJal;
    PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(True) then Exit ;
FValJal:=TFValJal.Create(Application) ;
if Validation then FValJal.HelpContext:=7700000 else FValJal.HelpContext:=7703000 ;
FValJal.Validation:=Validation ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FValJal.ShowModal ;
    Finally
     FValJal.Free ;
     _DeblocageMonoPoste(True) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FValJal,PP) ;
   FValJal.Show ;
   END ;
END ;

procedure TFValJal.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
CritModified := false;         {FP 10/11/2005 FQ14941}
end;

procedure TFValJal.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFValJal.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFValJal.FormShow(Sender: TObject);
begin

{$IFDEF EAGLCLIENT}
    BImprimer.Visible := false;
{$ENDIF}

    if Validation then Caption:=MsgBox.Mess[3] else Caption:=MsgBox.Mess[4] ;
    UpdateCaption(Self) ;
    FListe.ColWidths[FListe.ColCount-1]:=-1;
    RempliCbExo;
    ChercheJournaux;
    CbExo.Value := EXRF(VH^.Entree.Code);
    BChercheClick(Nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/09/2004
Modifié le ... : 22/09/2004
Description .. : - BPY le 22/09/2004 - optimization
Mots clefs ... : 
*****************************************************************}
procedure TFValJal.BChercheClick(Sender: TObject);
begin
    CritModified := false;         {FP 10/11/2005 FQ14941}
    // nettoyage de la grille des journaux
    FListe.ClearSelected;
    // nettoyage de la grille des periodes
    FListe1.ClearSelected;

    RempliLesDates;
    ExoValidOuiNon;
    AfficheResultat;
end;

{Procedure TFValJal.VideCombo(Sender : TComBoBox) ;
Var i : Byte ;
BEGIN
for i:=0 to TComBoBox(Sender).Items.Count-1 do
    if TComBoBox(Sender).Items.Objects[i]<>Nil then
       BEGIN
       TComBoBox(Sender).Items.Objects[i].Free ;
       TComBoBox(Sender).Items.Objects[i]:=Nil ;
       END ;
END ;}

procedure TFValJal.CbExoChange(Sender: TObject);
begin
    if (CbExo.Text = '') then Exit;
    CritModified := true;         {FP 10/11/2005 FQ14941}
    InitExo;

    BChercheClick(nil);
end;

Procedure TFValJal.InitExo ;
BEGIN
if CbExo.Value=VH^.Encours.Code then
   BEGIN
   Exo.Code:=VH^.Encours.Code ; Exo.Deb:=VH^.Encours.Deb ;
   Exo.Fin:=VH^.Encours.Fin ; NbrMois:=VH^.Encours.NombrePeriode ; PerioJal:='J_VALIDEEN' ;
   END else
   BEGIN
   Exo.Code:=VH^.Suivant.Code ; Exo.Deb:=VH^.Suivant.Deb ;
   Exo.Fin:=VH^.Suivant.Fin ; NbrMois:=VH^.Suivant.NombrePeriode ; PerioJal:='J_VALIDEEN1' ;
   END ;
END ;

Procedure TFValJal.RempliCbExo ;
Var Sql : String ;
    X : TInfoSup ;
    Q : TQuery;
BEGIN
CbExo.Values.Clear ; CbExo.Items.Clear ;
(*
if VH^.Suivant.Code<>'' then Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_DATEDEBUT>="'+USDateTime(VH^.EnCours.Deb)+'"'
                       else Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_EXERCICE="'+VH^.EnCours.Code+'"' ;
*)
if VH^.Suivant.Code<>'' then Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_EXERCICE="'+VH^.EnCours.Code+'" OR EX_EXERCICE="'+VH^.Suivant.Code+'"'
                        else Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_EXERCICE="'+VH^.EnCours.Code+'"' ;
Q := OpenSQL(Sql,True);
While Not Q.Eof do
   BEGIN
   X:=TInfoSup.Create ; X.St:=Q.Fields[2].AsString ;
   CbExo.Values.Add(Q.Fields[0].AsString) ; CbExo.Items.AddObJect(Q.Fields[1].AsString,X) ;
   Q.Next ;
   END ;
Ferme(Q);
END ;

Procedure TFValJal.RempliLesDates ;
Var i : Byte ;
    Dda,Fda : TDateTime ;
    StDate : String ;
BEGIN
CBDebDat.Items.Clear ; CBFinDate.Items.Clear ;
Dda:=Exo.Deb ; Fda:=FindeMois(Dda) ; Fliste1.RowCount:=NbrMois+1 ;
for i:=1 to NbrMois do
    BEGIN
    StDate:=FormatDateTime('mmmm-yyyy',Dda) ; StDate:=FirstMajuscule(StDate) ;
    FListe1.Cells[0,i]:=StDate ;
    CBDebDat.Items.Add(DatetoStr(Dda)) ; CBFinDate.Items.Add(DatetoStr(Fda)) ;
    Dda:=PlusMois(Dda,1) ; Fda:=FindeMois(Dda) ;
    END ;
Fliste1Click(Nil) ;
END ;

procedure TFValJal.Fliste1Click(Sender: TObject);
begin
PPer.Caption:='  '+MsgBox.Mess[2]+' '+FListe1.Cells[0,FListe1.Row] ;
FListe1.AllSelected := False;
AfficheResultat ; GeleLesBoutons ; CompteElemSelectionner ;
end;

Function TFValJal.FiltreJal : String ;
BEGIN
FiltreJal:=' J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" And J_NATUREJAL<>"ANA" ' ;
END ;

Procedure TFValJal.ChercheJournaux ;
Var Sql : String ;
    X,Y : TInfoSup ;
    Q : TQuery;
BEGIN
FListe.VidePile(False) ;
Sql:='Select J_JOURNAL,J_LIBELLE,J_VALIDEEN,J_VALIDEEN1 From JOURNAL Where '+FiltreJal+' Order by J_JOURNAL ' ;
Q := OpenSQL(Sql,True);
PerJal.Items.Clear ;
PerJal.Items.Clear ;
While Not Q.Eof do
  BEGIN
  X:=TInfoSup.Create ; X.St:=Q.Fields[2].AsString+';'+Q.Fields[1].AsString ;
  Y:=TInfoSup.Create ; Y.St:=Q.Fields[3].AsString+';'+Q.Fields[1].AsString ; ;
  PerJal.Items.AddObject(Q.Fields[0].AsString,X) ;
  PerJal1.Items.AddObject(Q.Fields[0].AsString,Y) ;
  Q.Next ;
  END ;
Ferme(Q);
END ;

Procedure TFValJal.ExoValidOuiNon ;
Var i,j : Byte ;
    Trouver : Boolean ;
BEGIN
for i:=1 to FListe1.RowCount-1 do
   BEGIN
   if (TInfoSup(CbExo.Items.Objects[CbExo.ItemIndex]).St[i])='X' then FListe1.Cells[1,i]:=MsgBox.Mess[7]
   else BEGIN
        Trouver:=False ;
        if PerioJal='J_VALIDEEN' then
           BEGIN
           for j:=0 to PerJal.Items.Count-1 do
               if TInfoSup(PerJal.Items.Objects[j]).St[i]='X' then BEGIN Trouver:=True ; Break ; END ;
           END else
           BEGIN
           for j:=0 to PerJal1.Items.Count-1 do
               if TInfoSup(PerJal1.Items.Objects[j]).St[i]='X' then BEGIN Trouver:=True ; Break ; END ;
           END ;
        if Trouver then FListe1.Cells[1,i]:=MsgBox.Mess[8]
                   else FListe1.Cells[1,i]:=MsgBox.Mess[9] ;
        END ;
   END ;
END ;

procedure TFValJal.FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if Not(ssCtrl in Shift) then Exit ;
if Button<>mbLeft then Exit ;
if Fliste.Cells[0,1]='' then Exit ;
if Fliste.Cells[FListe.ColCount-1,FListe.Row]='*'
   then Fliste.Cells[FListe.ColCount-1,FListe.Row]:=''
   else Fliste.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
FListe.Invalidate ; CompteElemSelectionner ;
end;

Procedure TFValJal.CompteElemSelectionner ;
Var i : Integer ;
BEGIN
TotalSelec:=0 ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(TotalSelec) ;
Nb1.Caption:=IntToStr(TotalSelec) ;
if TotalSelec>1 then Tex1.Caption:=MsgBox.Mess[6] else Tex1.Caption:=MsgBox.Mess[5] ;
END ;

procedure TFValJal.FlisteDblClick(Sender: TObject);
begin FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ; end;

Procedure TFValJal.AfficheResultat ;
Var i : Integer ; {JP 29/06/06 : FQ 18122 : si on a plus de 255 journaux, Integer est mieux que Byte}
    StC,StV : String ;
BEGIN
if PerJal.Items.Count<0 then Exit ;
if PerJal1.Items.Count<0 then Exit ;
FListe.VidePile(False) ; StC:='' ; StV:='' ;
Case CbVal.State of
     CbChecked :  BEGIN
                  if PerioJal='J_VALIDEEN' then
                     BEGIN
                     for i:=0 to PerJal.Items.Count-1 do
                         BEGIN
                         Stc:=TInfoSup(PerJal.Items.Objects[i]).St ;
                         StV:=ReadTokenSt(Stc) ;
                         if (Length(StV)<24) then StV := StV +  StringOfChar('-', 24-Length(StV)); // 14691 
                         if StV[FListe1.Row]='X' then
                            BEGIN
                            FListe.Cells[0,FListe.RowCount-1]:=PerJal.Items[i] ;
                            FListe.Cells[1,FListe.RowCount-1]:=Stc ;
                            FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[0] ;
                            FListe.RowCount:=FListe.RowCount+1 ;
                            END ;
                         END ;
                     END else
                     BEGIN
                     for i:=0 to PerJal1.Items.Count-1 do
                         BEGIN
                         Stc:=TInfoSup(PerJal1.Items.Objects[i]).St ;
                         StV:=ReadTokenSt(Stc) ;
                         if (Length(StV)<24) then StV := StV +  StringOfChar('-', 24-Length(StV)); // 14691 
                         if StV[FListe1.Row]='X' then
                            BEGIN
                            FListe.Cells[0,FListe.RowCount-1]:=PerJal.Items[i] ;
                            FListe.Cells[1,FListe.RowCount-1]:=Stc ;
                            FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[0] ;
                            FListe.RowCount:=FListe.RowCount+1 ;
                            END ;
                         END ;
                     END ;
                  END ;
     cbGrayed :   BEGIN
                  if PerioJal='J_VALIDEEN' then
                     BEGIN
                     for i:=0 to PerJal.Items.Count-1 do
                         BEGIN
                         Stc:=TInfoSup(PerJal.Items.Objects[i]).St ;
                         StV:=ReadTokenSt(Stc);
                         if (Length(StV)<24) then StV := StV +  StringOfChar('-', 24-Length(StV)); // 14691 
                         FListe.Cells[0,FListe.RowCount-1]:=PerJal.Items[i] ;
                         FListe.Cells[1,FListe.RowCount-1]:=Stc ;
                         if StV[FListe1.Row]='X' then FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[0]
                                                 else FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[1] ;
                         FListe.RowCount:=FListe.RowCount+1 ;
                         END ;
                     END else
                     BEGIN
                     for i:=0 to PerJal1.Items.Count-1 do
                         BEGIN
                         Stc:=TInfoSup(PerJal1.Items.Objects[i]).St ;
                         StV:=ReadTokenSt(Stc) ;
                         if (Length(StV)<24) then StV := StV +  StringOfChar('-', 24-Length(StV)); // 14691 
                         FListe.Cells[0,FListe.RowCount-1]:=PerJal.Items[i] ;
                         FListe.Cells[1,FListe.RowCount-1]:=Stc ;
                         if StV[FListe1.Row]='X' then FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[0]
                                                 else FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[1] ;
                         FListe.RowCount:=FListe.RowCount+1 ;
                         END ;
                     END ;
                  END ;
     CbUnChecked :BEGIN
                  if PerioJal='J_VALIDEEN' then
                     BEGIN
                     for i:=0 to PerJal.Items.Count-1 do
                         BEGIN
                         Stc:=TInfoSup(PerJal.Items.Objects[i]).St ;
                         StV:=ReadTokenSt(Stc) ;
                         if (Length(StV)<24) then StV := StV +  StringOfChar('-', 24-Length(StV)); // 14691 
                         if StV[FListe1.Row]='-' then
                            BEGIN
                            FListe.Cells[0,FListe.RowCount-1]:=PerJal.Items[i] ;
                            FListe.Cells[1,FListe.RowCount-1]:=Stc ;
                            FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[1] ;
                            FListe.RowCount:=FListe.RowCount+1 ;
                            END ;
                         END ;
                     END else
                     BEGIN
                     for i:=0 to PerJal1.Items.Count-1 do
                         BEGIN
                         Stc:=TInfoSup(PerJal1.Items.Objects[i]).St ;
                         StV:=ReadTokenSt(Stc) ;
                         if (Length(StV)<24) then StV := StV +  StringOfChar('-', 24-Length(StV)); // 14691 
                         if StV[FListe1.Row]='-' then
                            BEGIN
                            FListe.Cells[0,FListe.RowCount-1]:=PerJal.Items[i] ;
                            FListe.Cells[1,FListe.RowCount-1]:=Stc ;
                            FListe.Cells[2,FListe.RowCount-1]:=MsgBox.Mess[1] ;
                            FListe.RowCount:=FListe.RowCount+1 ;
                            END ;
                         END ;
                     END ;
                  END ;
     End ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
BTag.Visible:=True ; Bdetag.Visible:=False ;
END ;

procedure TFValJal.BFirstClick(Sender: TObject);
begin FListe1.Row:=1 ; GeleLesBoutons ; end;

procedure TFValJal.BPrevClick(Sender: TObject);
begin FListe1.Row:=FListe1.Row-1 ; GeleLesBoutons ; end;

procedure TFValJal.BNextClick(Sender: TObject);
begin FListe1.Row:=FListe1.Row+1 ; GeleLesBoutons ; end;

procedure TFValJal.BLastClick(Sender: TObject);
begin FListe1.Row:=FListe1.RowCount-1 ; GeleLesBoutons ; end;

Procedure TFValJal.GeleLesBoutons ;
BEGIN
BFirst.Enabled:=(FListe1.Row>1) ; BPrev.Enabled:=(FListe1.Row>1) ;
BNext.Enabled:=(FListe1.Row<FListe1.RowCount-1) ; BLast.Enabled:=(FListe1.Row<FListe1.RowCount-1) ;
END ;

Procedure TFValJal.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
if Fliste.Cells[0,1]='' then Exit ;
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ;
end;

procedure TFValJal.BTagClick(Sender: TObject);
begin TagDetag(True) ; CompteElemSelectionner ; end;

procedure TFValJal.BdetagClick(Sender: TObject);
begin TagDetag(False) ; CompteElemSelectionner ; end;

procedure TFValJal.BValiderClick(Sender: TObject);
begin
  NbPerio:=FListe1.Row ;
  if Not OnExecute   then Exit ;
  if Validation then
  begin
    If MsgBox.Execute(11,caption,'')=mrYes then MajDesTables;
  end
  else
  begin
{$IFNDEF CERTIFNF}
    if GetParamSocSecur('SO_CPCONFORMEBOI', False) then  //06/12/2006 YMO Norme NF 203
    begin
      PgiInfo('Pour la conformité stricte avec la norme NF 203 (et le BOI du 24/01/2006) cette fonction n''est plus disponible',Caption);
      Exit;
    end;
{$ENDIF}

    if MsgBox.Execute(12,caption,'')=mrYes then
    begin
      PgiInfo('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 23 qui rappelle l''interdiction de ' + #13#10 +
              'toute modification et/ou suppression d''écriture validée, nous vous conseillons de « tracer » l’information ' + #13#10 +
              'par toute méthode à votre disposition (par ex le bloc note).', Caption);
      MajDesTables;
    end;
  end;
  BChercheClick(Nil) ;
end;

Function TFValJal.OnExecute : Boolean ;
Var i : Byte ;
    AFaire : Boolean ;
    St : String ;
BEGIN
AFaire:=False ; Result:=True ;
St:=TInfoSup(CbExo.Items.Objects[CbExo.ItemIndex]).St ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then BEGIN AFaire:=True ; Break ; END ;
if Not AFaire then BEGIN Result:=False ; MsgBox.Execute(13,caption,'') ; Exit ; END ;
AFaire:=False ;
if Validation then
   BEGIN
   for i:=1 to FListe.RowCount-1 do
       if FListe.Cells[FListe.ColCount-1,i]='*' then
          if FListe.Cells[2,i]=MsgBox.Mess[1] then BEGIN AFaire:=True ; Break ; END ;
   if Not AFaire then BEGIN Result:=False ; MsgBox.Execute(10,caption,'') ; Exit ; END ;
   END else
   BEGIN
   for i:=1 to FListe.RowCount-1 do
       if FListe.Cells[FListe.ColCount-1,i]='*' then
          if FListe.Cells[2,i]=MsgBox.Mess[0] then BEGIN AFaire:=True ; Break ; END ;
   if Not AFaire then BEGIN Result:=False ; MsgBox.Execute(10,caption,'') ; Exit ; END ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/09/2004
Modifié le ... : 07/12/2006
Description .. : - BPY 22/09/2004 - optimisation suite a la fiche de bug n°
Description .. : - YMO 07/12/2006 - Norme NF : Ecriture dans le Log pour chaque écriture
Suite ........ : 14635
Mots clefs ... : 
*****************************************************************}
Procedure TFValJal.MajDesTables;
Var
    Sql,St : String;
    i : Byte;
{$IFNDEF CERTIFNF}
    j : integer;
{$ENDIF}
    DateD,DateF : String;
    ListePiece : HTStringList;
    QPiece : TQuery;
begin
    ListePiece:=HTStringList.Create ;
    ListePiece.Clear ;
    ListePiece.Sorted:=TRUE ;

{$IFDEF CERTIFNF}
    If Validation Then
       ListePiece.Add('VALECR ')
    Else
       ListePiece.Add('ANNVALECR ');
{$ELSE}
    ListePiece.Sorted:=TRUE ;
{$ENDIF}

    if (Validation) then            // VALIDATION
    begin
        // init des dates
        DateD := UsDateTime(StrToDate(CBDebDat.Items[0]));
        DateF := UsDateTime(StrToDate(CBFinDate.Items[NbPerio-1]));

        for i := 1 to FListe.RowCount-1 do // pour chaque journal
        begin
            if ((FListe.Cells[FListe.ColCount-1,i] = '*') And (FListe.Cells[2,i] = MsgBox.Mess[1])) then
            begin
                // mise a jour de la variable de validation
                st := MajObjetJalVal(FListe.Cells[0,i]);
                if (St = '') then Continue;

                // mise a jour du journal
                ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + St + '" WHERE J_JOURNAL="' + FListe.Cells[0,i] + '"');

                // mise a jour des ecritures
                Sql := 'UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + CbExo.Value + '" AND E_DATECOMPTABLE>="'
                     + DateD + '" AND E_DATECOMPTABLE<="' + DateF + '" AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I") AND ' // SBO 19/08/2004 FQ14209 prise en compte des écritures "I"
                     + 'E_ECRANOUVEAU="N" And E_JOURNAL="' + FListe.Cells[0,i] + '"';
                ExecuteSQL(Sql);

                QPiece:=OpenSQL(StringReplace(Sql,'UPDATE ECRITURE SET E_VALIDE="X"','SELECT E_JOURNAL,E_PERIODE,E_NUMEROPIECE FROM ECRITURE',[]),True) ;
                While Not QPiece.EOF do
                BEGIN
                   ListePiece.Add(QPiece.Fields[0].AsString+'-'+QPiece.Fields[1].AsString+'-'+QPiece.Fields[2].AsString);
                   QPiece.Next ;
                END ;
                Ferme(QPiece) ;

                // mise a jour des ventilations // FQ16762 SBO 04/10/2005 Prise en compte des ventilations à la validation
                Sql := 'UPDATE ANALYTIQ SET Y_VALIDE="X" WHERE Y_EXERCICE="' + CbExo.Value + '" AND Y_DATECOMPTABLE>="' + DateD
                     + '" AND Y_DATECOMPTABLE<="' + DateF + '" AND (Y_QUALIFPIECE="N" OR Y_QUALIFPIECE="I") '
                     + 'AND Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N" ';
                ExecuteSQL(Sql);

                // GCO - 18/09/2006
                // YMO 07/12/2006 Norme NF Suppression
                // CPEnregistreLog('VALPERIODE ' + FListe.Cells[0,i] + ' - ' + FormatDateTime('mmm yyyy', StrToDate(CBFinDate.Items[NbPerio-1])));

            end;
        end;

        // validation de la periode dans l'exercice
        if (OnValideAussiLexo) then
        begin
            St := TInfoSup(CbExo.Items.Objects[CbExo.ItemIndex]).St;
            for i := 1 to NbPerio do St[i] := 'X';
            TInfoSup(CbExo.Items.Objects[CbExo.ItemIndex]).St := St;
            ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + St + '" WHERE EX_EXERCICE="' + CbExo.Value + '"');
        end;
    end
    else                            // DEVALIDATION
    begin
        // test : on ne devalide pas une periode cloturé !
        if (StrToDate(CBDebDat.Items[NbPerio-1]) <= VH^.DateCloturePer) then
        begin
            MsgBox.Execute(14,Caption,'');
            Exit;
        end;

        // init des dates
        // DateD := UsDateTime(StrToDate(CBDebDat.Items[NbPerio-1]));
        { FQ 16139 - CA - 11/07/2005 : Si date début d'exercice <> 1er jour du mois, comparer par rapport
        au premier jour du mois à partir du deuxième mois de l'exercice }
        if (NbPerio > 1) then DateD := UsDateTime(DebutDeMois(StrToDate(CBDebDat.Items[NbPerio-1])))
        else  DateD := UsDateTime(StrToDate(CBDebDat.Items[NbPerio-1]));
        DateF := UsDateTime(StrToDate(CBFinDate.Items[NbPerio-1]));

        // devalidation de la periode dans l'exercice
        St := TInfoSup(CbExo.Items.Objects[CbExo.ItemIndex]).St;
        if (St[NbPerio] = 'X') then
        begin
            St[NbPerio] := '-';
            TInfoSup(CbExo.Items.Objects[CbExo.ItemIndex]).St := St;
            ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + St + '" WHERE EX_EXERCICE="' + CbExo.Value + '"');
        end;

        for i := 1 to FListe.RowCount-1 do // pour chaque journal
        begin
            if (FListe.Cells[FListe.ColCount-1,i] = '*') And (FListe.Cells[2,i] = MsgBox.Mess[0]) then
            begin
                // mise a jour de la variable de validation
                St := MajObjetJalDeVal(FListe.Cells[0,i]);
                if St = '' then Continue;

                // mise a jour du journal
                ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + St + '" WHERE J_JOURNAL="' + FListe.Cells[0,i] + '"');

                // mise a jour des ecriture
                Sql := 'UPDATE ECRITURE SET E_VALIDE="-" WHERE E_EXERCICE="' + CbExo.Value + '" AND E_DATECOMPTABLE>="'
                     + DateD + '" AND E_DATECOMPTABLE<="' + DateF + '" AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I") AND ' // SBO 19/08/2004 FQ14209 prise en compte des écritures "I"
                     + 'E_ECRANOUVEAU="N"  AND E_CREERPAR<>"DET" AND E_JOURNAL="' + FListe.Cells[0,i] + '"';
                { CA - 10/10/2003 - FQ 12448 - On ne touche pas aux pièces validées par la gescom
                  JP - 29/07/05 : FQ 15124 : On revient en arrière
                Sql := Sql + ' AND E_REFGESCOM=""';
                 JP 26/06/07 : Par contre on ne touche pas aux pièces de Tréso}
                Sql := Sql + ' AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL OR E_QUALIFORIGINE = "")';
                ExecuteSQL(Sql);

                QPiece:=OpenSQL(StringReplace(Sql,'UPDATE ECRITURE SET E_VALIDE="-"','SELECT E_JOURNAL,E_PERIODE,E_NUMEROPIECE FROM ECRITURE',[]),True) ;
                While Not QPiece.EOF do
                BEGIN
                   ListePiece.Add(QPiece.Fields[0].AsString+'-'+QPiece.Fields[1].AsString+'-'+QPiece.Fields[2].AsString);
                   QPiece.Next ;
                END ;
                Ferme(QPiece) ;

                // mise a jour des ventilations // FQ16762 SBO 04/10/2005 Prise en compte des ventilations à la validation
                Sql := 'UPDATE ANALYTIQ SET Y_VALIDE="-" WHERE Y_EXERCICE="' + CbExo.Value + '" AND Y_DATECOMPTABLE>="' + DateD
                     + '" AND Y_DATECOMPTABLE<="' + DateF + '" AND (Y_QUALIFPIECE="N" OR Y_QUALIFPIECE="I") '
                     + 'AND Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N" ';
                ExecuteSQL(Sql);

                // GCO - 18/09/2006
                // YMO 07/12/2006 Norme NF Suppression
                //  CPEnregistreLog('ANNULVALPERIODE ' + FListe.Cells[0,i] + ' - ' + FormatDateTime('mmm yyyy', StrToDate(CBFinDate.Items[NbPerio-1])));

            end;
        end;
    end;
    
    { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
    For j:=0 To ListePiece.Count-1 Do //YMO 09/01/2007 Si on prend un Byte, ça passe dans la boucle de 0 à -1(count=0)
    Begin
     If Validation Then
        CPEnregistreLog('VALECR ' + ListePiece[j])
     Else
        CPEnregistreLog('ANNVALECR ' + ListePiece[j]);
    End;
{$ELSE}
    CPEnregistreJalEvent('CVE','Validation d''écritures',ListePiece);
{$ENDIF}

    ListePiece.Free ;

    // recherche ....
    BChercheClick(Nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/09/2004
Modifié le ... : 22/09/2004
Description .. : - BPY 22/09/2004 - optimisation suite a la fiche de bug n°
Suite ........ : 14635
Mots clefs ... :
*****************************************************************}
function TFValJal.MajObjetJalVal(StCod : String) : string;
Var
    i : Byte;
    St : String;
begin
    result := '';
    if (StCod = '') then Exit;

    if (PerioJal = 'J_VALIDEEN') then
    begin
        St := TInfoSup(PerJal.Items.Objects[PerJal.Items.Indexof(StCod)]).St;
        for i := 1 to NbPerio do St[i] := 'X';
        TInfoSup(PerJal.Items.Objects[PerJal.Items.Indexof(StCod)]).St := St;
    end
    else
    begin
        St := TInfoSup(PerJal1.Items.Objects[PerJal1.Items.Indexof(StCod)]).St;
        for i := 1 to NbPerio do St[i] := 'X';
        TInfoSup(PerJal1.Items.Objects[PerJal1.Items.Indexof(StCod)]).St := St;
    end;

    Result := ReadTokenSt(St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/09/2004
Modifié le ... : 22/09/2004
Description .. : - BPY 22/09/2004 - optimisation suite a la fiche de bug n° 
Suite ........ : 14635
Mots clefs ... : 
*****************************************************************}
function TFValJal.MajObjetJalDeVal(StCod : String) : String;
Var
    St : String;
begin
    Result := '';
    if (StCod = '') then Exit;

    if (PerioJal = 'J_VALIDEEN') then
    begin
        TinfoSup(PerJal.Items.Objects[PerJal.Items.IndexOf(StCod)]).St[NbPerio] := '-';
        St := TinfoSup(PerJal.Items.Objects[PerJal.Items.IndexOf(StCod)]).St;
    end
    else
    begin
        TinfoSup(PerJal1.Items.Objects[PerJal1.Items.IndexOf(StCod)]).St[NbPerio] := '-';
        St := TinfoSup(PerJal1.Items.Objects[PerJal1.Items.IndexOf(StCod)]).St;
    end;

    Result := ReadTokenSt(St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/09/2004
Modifié le ... : 22/09/2004
Description .. : - BPY 22/09/2004 - optimisation suite a la fiche de bug n° 
Suite ........ : 14635
Mots clefs ... : 
*****************************************************************}
Function TFValJal.OnValideAussiLexo : Boolean;
Var
    i : Byte;
    Trouver : Boolean;
begin
    Trouver:=False ;

    if (PerioJal = 'J_VALIDEEN') then
    begin
        for i := 0 to PerJal.Items.Count-1 do
            if (TInfoSup(PerJal.Items.Objects[i]).St[NbPerio] = '-') then
            begin
                Trouver := True;
                Break;
            end;
    end
    else
    begin
        for i := 0 to PerJal1.Items.Count-1 do
            if (TInfoSup(PerJal1.Items.Objects[i]).St[NbPerio] = '-') then
            begin
                Trouver := True;
                Break;
            end;
    end;
    Result := (Not Trouver);
end;

procedure TFValJal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFValJal.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
// A FAIRE
{$ELSE}
PrintDBGrid(FListe,PExo,Caption,'') ;
{$ENDIF}
end;

procedure TFValJal.FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((Shift=[]) And (Key=VK_SPACE)) or ((Shift=[ssShift]) And (Key=VK_DOWN)) then
   FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ;
end;

procedure TFValJal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ; FListe1.VidePile(True) ;
if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True) ;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    Action:=caFree ;
  {$ENDIF}
   END ;
end;

procedure TFValJal.SetCritModified(const Value: Boolean);
begin
  {b FP 10/11/2005 FQ14941}
  if FCritModified <> Value then
    begin
    if Value then
      BCherche.Glyph := iCritGlyphModified.Picture.BitMap
    else
      BCherche.Glyph := iCritGlyph.Picture.BitMap;
    end;
  FCritModified := Value;
  {e FP 10/11/2005 FQ14941}
end;

procedure TFValJal.CbValClick(Sender: TObject);
begin
CritModified := true;         {FP 10/11/2005 FQ14941}
end;

end.
