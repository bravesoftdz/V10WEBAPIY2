{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit GenToBud;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, Menus, HSysMenu, StdCtrls, Hctrls, HTB97, Buttons, ExtCtrls,
  Grids, ComCtrls, Ent1, HEnt1, Hcompte,
{$IFDEF EAGLCLIENT}

{$ELSE}
  Db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HStatus, MajCodBu, HPanel, UiUtil, UTob,
  UObjFiltres; {SG6 04/01/05 Gestion Filtres V6 FQ 15145}

Procedure CodeGeneVersCodeBudg ;

Type TInfoGene = Class
       Cpte : String ;
       Lib : String ;
       Sens : String ;
       User : String ;
     End ;  

type
  TFGenToBud = class(TForm)
    Pages: TPageControl;
    Pparam: TTabSheet;
    Bevel1: TBevel;
    FListe: THGrid;
    HPB: TToolWindow97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BCherche: TToolbarButton97;
    HM: THMsgBox;
    TG_NATUREGENE: THLabel;
    G_NATUREGENE: THValComboBox;
    TG_GENERAL: TLabel;
    G_GENERAL: THCpteEdit;
    TG_GENERAL_: TLabel;
    G_GENERAL_: THCpteEdit;
    TG_VENTILABLE: THLabel;
    G_VENTILABLE: THValComboBox;
    PzLibre: TTabSheet;
    Bevel5: TBevel;
    TG_TABLE0: TLabel;
    TG_TABLE1: TLabel;
    TG_TABLE2: TLabel;
    TG_TABLE3: TLabel;
    TG_TABLE4: TLabel;
    TG_TABLE5: TLabel;
    TG_TABLE6: TLabel;
    TG_TABLE7: TLabel;
    TG_TABLE8: TLabel;
    TG_TABLE9: TLabel;
    G_TABLE0: THCpteEdit;
    G_TABLE1: THCpteEdit;
    G_TABLE2: THCpteEdit;
    G_TABLE3: THCpteEdit;
    G_TABLE4: THCpteEdit;
    G_TABLE5: THCpteEdit;
    G_TABLE6: THCpteEdit;
    G_TABLE7: THCpteEdit;
    G_TABLE8: THCpteEdit;
    G_TABLE9: THCpteEdit;
    Nb1: TLabel;
    Tex1: TLabel;
    Bdetag: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BTag: TToolbarButton97;
    Dock: TDock97;
    Dock971: TDock97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure G_NATUREGENEChange(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    //SG6 04/01/05 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    
    WMinx,WMiny : Integer ;
    TempT : TDateTime ;
    Procedure GetMinMaxInfo(var M : TMessage) ; message WM_GETMINMAXINFO ;
    Procedure FirstLastHcpteGen(UnZoom : TZoomTable ; Var C1,C2 : String) ;
    Function  FaitRequete : String ;
    Function  CreerUnObj(Q : TQuery) : TInfoGene ;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure CompteElemSelectionner ;
    Function  ListeVide : Boolean ;
    Procedure InverseSelection ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Function  CodeRubriqueExiste(St : String) : Boolean ;
    Function  TrouveCle(St : String) : Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses
  ed_tools, UObjEtats ;

Procedure CodeGeneVersCodeBudg ;
var FGenToBud: TFGenToBud;
    PP : THPanel ;
BEGIN
FGenToBud:=TFGenToBud.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FGenToBud.ShowModal ;
    Finally
     FGenToBud.Free ;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FGenToBud,PP) ;
   FGenToBud.Show ;
   END ;
END ;

procedure TFGenToBud.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'GENTOBUD');
WMinx:=Width ; WMiny:=Height ;
end;

Procedure TFGenToBud.GetMinMaxInfo(Var M : TMessage) ;
BEGIN
with PMinMaxInfo(M.lparam)^.ptMinTrackSize do begin X:=WMinX ; Y:=WMinY ; end ;
END ;

procedure TFGenToBud.FormShow(Sender: TObject);
begin
FListe.GetCellCanvas:=GetCellCanvas ;
G_NATUREGENE.ItemIndex:=0 ; G_VENTILABLE.ItemIndex:=0 ;
G_NATUREGENEChange(Nil) ;
//SG6 04/01/05 Gestion Filtrer V6 FQ 15145
ObjFiltre.Charger;
end;

Procedure TFGenToBud.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFGenToBud.G_NATUREGENEChange(Sender: TObject);
Var St,C1,C2 : String ;
begin
St:=G_NATUREGENE.Value ;
if St=''    then BEGIN G_GENERAL.ZoomTable:=tzGeneral ; G_GENERAL_.ZoomTable:=tzGeneral ; END else
if St='BQE' then BEGIN G_GENERAL.ZoomTable:=tzGBanque ; G_GENERAL_.ZoomTable:=tzGBanque ; END else
if St='CAI' then BEGIN G_GENERAL.ZoomTable:=tzGCaisse ; G_GENERAL_.ZoomTable:=tzGCaisse ; END else
if St='CHA' then BEGIN G_GENERAL.ZoomTable:=tzGCharge ; G_GENERAL_.ZoomTable:=tzGCharge ; END else
if St='COC' then BEGIN G_GENERAL.ZoomTable:=tzGCollClient ; G_GENERAL_.ZoomTable:=tzGCollClient ; END else
if St='COD' then BEGIN G_GENERAL.ZoomTable:=tzGCollDivers ; G_GENERAL_.ZoomTable:=tzGCollDivers ; END else
if St='COF' then BEGIN G_GENERAL.ZoomTable:=tzGCollFourn ; G_GENERAL_.ZoomTable:=tzGCollFourn ; END else
if St='COS' then BEGIN G_GENERAL.ZoomTable:=tzGCollSalarie ; G_GENERAL_.ZoomTable:=tzGCollSalarie ; END else
if St='DIV' then BEGIN G_GENERAL.ZoomTable:=tzGDivers ; G_GENERAL_.ZoomTable:=tzGDivers ; END else
if St='EXT' then BEGIN G_GENERAL.ZoomTable:=tzGExtra ; G_GENERAL_.ZoomTable:=tzGExtra ; END else
if St='IMO' then BEGIN G_GENERAL.ZoomTable:=tzGImmo ; G_GENERAL_.ZoomTable:=tzGImmo ; END else
if St='PRO' then BEGIN G_GENERAL.ZoomTable:=tzGProduit ; G_GENERAL_.ZoomTable:=tzGProduit ; END else
if St='TIC' then BEGIN G_GENERAL.ZoomTable:=tzGTIC ; G_GENERAL_.ZoomTable:=tzGTIC ; END else
if St='TID' then BEGIN G_GENERAL.ZoomTable:=tzGTID ; G_GENERAL_.ZoomTable:=tzGTID ; END ;
FirstLastHcpteGen(G_GENERAL.ZoomTable,C1,C2) ;
G_GENERAL.Text:=C1 ; G_GENERAL_.Text:=C2 ;
end;

Function SqlFirstLastHcpteGen(UnZoom : TZoomTable ; Premier : Boolean) : String ;
Var Sql,SqlOrder,SqlWhere : String ;
    QLoc : TQuery ;
BEGIN
Sql:='Select G_GENERAL From GENERAUX ' ;
if Premier then SqlOrder:='Order by G_GENERAL ASC'
           else SqlOrder:='Order by G_GENERAL DESC' ;
Case UnZoom of
     tzGCollClient  : SqlWhere:='Where G_NATUREGENE="COC" ' ;
     tzGCollFourn   : SqlWhere:='Where G_NATUREGENE="COF" ' ;
     tzGCollDivers  : SqlWhere:='Where G_NATUREGENE="COD" ' ;
     tzGCollSalarie : SqlWhere:='Where G_NATUREGENE="COS" ' ;
     tzGTID         : SqlWhere:='Where G_NATUREGENE="TID" ' ;
     tzGTIC         : SqlWhere:='Where G_NATUREGENE="TIC" ' ;
     tzGBanque      : SqlWhere:='Where G_NATUREGENE="BQE" ' ;
     tzGCaisse      : SqlWhere:='Where G_NATUREGENE="CAI" ' ;
     tzGCharge      : SqlWhere:='Where G_NATUREGENE="CHA" ' ;
     tzGProduit     : SqlWhere:='Where G_NATUREGENE="PRO" ' ;
     tzGImmo        : SqlWhere:='Where G_NATUREGENE="IMO" ' ;
     tzGDivers      : SqlWhere:='Where G_NATUREGENE="DIV" ' ;
     tzGExtra       : SqlWhere:='Where G_NATUREGENE="EXT" ' ;
     else SqlWhere:='' ;
  End ;
Sql:=Sql+SqlWhere+SqlOrder ; QLoc:=OpenSql(Sql,True) ; Result:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
END ;

Procedure TFGenToBud.FirstLastHcpteGen(UnZoom : TZoomTable ; Var C1,C2 : String) ;
BEGIN
C1:=SqlFirstLastHcpteGen(UnZoom,True) ;
C2:=SqlFirstLastHcpteGen(UnZoom,False) ;
END ;

Function TFGenToBud.CreerUnObj(Q : TQuery) : TInfoGene ;
Var X : TInfoGene ;
BEGIN
X:=TInfoGene.Create ;
X.Cpte:=Q.FindField('G_GENERAL').AsString ;
X.Lib:=Q.FindField('G_LIBELLE').AsString ;
X.Sens:=Q.FindField('G_SENS').AsString ;
X.User:=Q.FindField('G_UTILISATEUR').AsString ;
Result:=X ;
END ;

procedure TFGenToBud.BChercheClick(Sender: TObject);
Var QLoc : TQuery ;
    SaveGestionMessage : Boolean ;
    X : TInfoGene ;
begin
FListe.VidePile(False) ;
SaveGestionMessage:=V_PGI.GestionMessage ; V_PGI.GestionMessage:=False ;
if (G_GENERAL.Text='') or (G_GENERAL_.Text='') then  BEGIN HM.Execute(2,'','') ; Exit ; END ;
QLoc:=OpenSql(FaitRequete,True) ;
While Not QLoc.Eof do
   BEGIN
   X:=CreerUnObj(QLoc) ;
   FListe.Cells[0,FListe.RowCount-1]:=X.Cpte ;
   FListe.Cells[1,FListe.RowCount-1]:=X.Lib ;
   FListe.Cells[2,FListe.RowCount-1]:=RechDom('TTSENS',X.Sens,False) ;
   FListe.Cells[3,FListe.RowCount-1]:=RechDom('TTUTILISATEUR',X.User,False) ;
   FListe.Objects[0,FListe.RowCount-1]:=X ;
   FListe.RowCount:=FListe.RowCount+1 ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
V_PGI.GestionMessage:=SaveGestionMessage ; BTagClick(Nil) ;
end;

Function TFGenToBud.FaitRequete : String ;
Var Sql,St : String ;
    i : Integer ;
BEGIN
Sql:='Select * From GENERAUX Where G_GENERAL>="'+G_GENERAL.Text+'" And G_GENERAL<="'+G_GENERAL_.Text+'" ' ;
if G_VENTILABLE.Value<>'' then
   BEGIN
   Case G_VENTILABLE.Value[2] of
        '1' : St:= 'And G_VENTILABLE1="X"' ;
        '2' : St:= 'And G_VENTILABLE2="X"' ;
        '3' : St:= 'And G_VENTILABLE3="X"' ;
        '4' : St:= 'And G_VENTILABLE4="X"' ;
        '5' : St:= 'And G_VENTILABLE5="X"' ;
      End ;
   Sql:=Sql+'And G_VENTILABLE="X" '+St ;
   END ;
if PZLibre.Visible then
   BEGIN
   St:='' ;
   for i:=0 to 9 do
       BEGIN
       if (THCpteEdit(FindComponent('G_TABLE'+IntToStr(i))).Enabled) And
          (THCpteEdit(FindComponent('G_TABLE'+IntToStr(i))).Text<>'')then
          St:=St+'And G_TABLE'+IntToStr(i)+'="'+THCpteEdit(FindComponent('G_TABLE'+IntToStr(i))).Text+'" ' ;
       END ;
   Sql:=Sql+St ;
   END ;
Result:=Sql ;
END ;

procedure TFGenToBud.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ;
//SG6 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Free;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFGenToBud.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

procedure TFGenToBud.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

Procedure TFGenToBud.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; FListe.SetFocus ;
Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ; CompteElemSelectionner ;
end;

Procedure TFGenToBud.CompteElemSelectionner ;
Var i,j : Integer ;
BEGIN
j:=0 ;
if Not ListeVide then for i:=1 to FListe.RowCount-1 do if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(j) ;
Case j of
     0,1: BEGIN Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[0] ; END ;
     else BEGIN Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[1] ; END ;
   End ;
END ;

Function TFGenToBud.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

Procedure TFGenToBud.InverseSelection ;
BEGIN
if ListeVide then Exit ;
if FListe.Cells[FListe.ColCount-1,FListe.Row]='*' then FListe.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                  else FListe.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
CompteElemSelectionner ; FListe.Invalidate ;
END ;

procedure TFGenToBud.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin if (ssCtrl in Shift) And (Button=mbLeft)then InverseSelection ; end;

procedure TFGenToBud.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((ssShift in Shift) And (Key=VK_DOWN)) or ((ssShift in Shift) And (Key=VK_UP)) then InverseSelection else
   if (Shift=[]) And (Key=VK_SPACE) then
      BEGIN
      InverseSelection ;
      if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
      END ;
end;

procedure TFGenToBud.HMTradBeforeTraduc(Sender: TObject);
begin
inherited;
LibellesTableLibre(PzLibre,'TG_TABLE','G_TABLE','G') ;
end;

Function TFGenToBud.CodeRubriqueExiste(St : String) : Boolean ;
BEGIN
  Result := ExisteSQL('SELECT BG_BUDGENE,BG_RUB FROM BUDGENE WHERE BG_BUDGENE<>"'+St+'" AND BG_RUB="'+Copy(St,1,5)+'"');
END ;

Function TFGenToBud.TrouveCle(St : String) : Boolean ;
BEGIN
  Result := ExisteSQL('SELECT * FROM BUDGENE WHERE BG_BUDGENE="'+St+'" AND BG_RUB="'+Copy(St,1,5)+'"');
END ;

procedure TFGenToBud.BValiderClick(Sender: TObject);
Var
  i : Integer ;
  Compteur : Integer ;
  Doublons : TOB;
  Gene : TInfoGene;
  T : Tob;
  F : TOB;
begin
  if ListeVide then BEGIN HM.Execute(3,'','') ; Exit ; END ;
  Compteur:=0 ;
  for i:=1 to FListe.RowCount-1 do if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(Compteur) ;
  if Compteur=0 then BEGIN HM.Execute(3,'','') ; Exit ; END ;
  if HM.Execute(4,'','') <> mrYes then Exit ;

  Doublons := TOB.Create('_DOUBLONS', nil, -1);
  TempT:=Date ;
  try
    InitMove(Compteur,'') ;
    T := Tob.Create('BUDGENE', nil, -1);
    try
      for i:=1 to FListe.RowCount-1 do
        BEGIN
        if FListe.Cells[FListe.ColCount-1,i]<>'*' then Continue ;
        MoveCur(False) ;
        if CodeRubriqueExiste(FListe.Cells[0,i]) then begin
          F := TOB.Create('***', Doublons, -1);
          Gene := TInfoGene(FListe.Objects[0, i]);
          F.AddChampSupValeur('COMPTE', Gene.Lib);
          F.AddChampSupValeur('CODE', Gene.Cpte);
          F.AddChampSupValeur('RUBRIQUE', Copy(Gene.Cpte, 1, 5));
          Continue;
        end;

        Gene := TInfoGene(FListe.Objects[0,i]);
        if TrouveCle(FListe.Cells[0,i]) then ExecuteSQL('DELETE FROM BUDGENE WHERE BG_BUDGENE="'+Gene.Cpte+'"');

        T.InitValeurs(False);
        T.SetString('BG_BUDGENE', Gene.Cpte);
        T.SetString('BG_LIBELLE', Gene.Lib);
        T.SetString('BG_ABREGE', Copy(Gene.Lib,1,17));
        T.SetString('BG_COMPTERUB', Gene.Cpte+';;');
        T.SetDateTime('BG_DATECREATION', TempT);
        T.SetDateTime('BG_DATEMODIF', TempT);
        T.SetString('BG_SENS', Gene.Sens);
        T.SetString('BG_UTILISATEUR', Gene.User);
        T.SetString('BG_RUB', Copy(Gene.Cpte,1,5));
        T.InsertOrUpdateDB;
        END ;
    finally
      FiniMove ;
      T.Free;
    end;

  {JP 23/04/07 : FQ 19369 : seul les 5 premiers caractères sont significatifs dans le Buget : ainsi
                 si on a les généraux suivants 601010, 601011, 601012, un seul compte budgétaire sera
                 créé. Voir la fonction CodeRubriqueExiste dans l'unité}
  if Doublons.Detail.Count > 0 then begin
    if PGIAsk(TraduireMemoire('Seuls les cinq premiers caractères des comptes généraux sont signgificatifs ;') + #13 +
              TraduireMemoire('de ce fait un certain nombres de comptes n''ont pas été pris en compte.') + #13#13 +
              TraduireMemoire('Voulez-vous en voir le détail ?'), Caption) = mrYes then begin
      TObjEtats.GenereEtatTob(Doublons, TraduireMemoire('Liste des comptes ignorés')); 

    end;
  end;

  finally
    FreeAndNil(Doublons);
  end;
end;

procedure TFGenToBud.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFGenToBud.BFermeClick(Sender: TObject);
begin
  //SG6 04/01/05 Vide le panel
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFGenToBud.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //SG6 04/01/05 Gestion des Filtres V6 FQ 15145
  if Key = VK_F9 then BChercheClick(nil);
end;

end.
