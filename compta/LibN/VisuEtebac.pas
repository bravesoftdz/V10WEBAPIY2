{***********UNITE*************************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/05/2002
Modifié le ... :   /  /    
Description .. : Mul de recherche des trous dans les numero de cheque ( 
Suite ........ : champ CRL_REFINTERNE dans la table CRLEBQE )
Mots clefs ... : 
*****************************************************************}
unit VisuEtebac;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, HTB97, ExtCtrls, StdCtrls, Mask, ColMemo, ComCtrls ,
  {$IFDEF EAGLCLIENT}
  eFiche,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  PrintDBG, Fiche,
  {$ENDIF EAGLCLIENT}
  UTOB,
  HEnt1,
  HPanel,
  UObjFiltres,  {JP 04/10/04 : Gestion des filtres}
  UIUtil, HSysMenu, Menus;

type
  TFVisuEtebac = class(TForm)
    Pages: TPageControl;
    PCritere: TTabSheet;
    Bevel1: TBevel;
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    Dock972: TDock97;
    PanelBouton: TToolWindow97;
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    bSelectAll: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BParamListe: TToolbarButton97;
    bExport: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BOuvrir: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    Binsert: TToolbarButton97;
    BBlocNote: TToolbarButton97;
    FListe: THGrid;
    HLabel1: THLabel;
    EdtJournal: THCritMaskEdit;
    HLabel2: THLabel;
    edtMulDateDu: THCritMaskEdit;
    HLabel3: THLabel;
    EdtMulDateAu: THCritMaskEdit;
    edtReference: THCritMaskEdit;
    HLabel4: THLabel;
    HLabel5: THLabel;
    EdtAmplitude: THCritMaskEdit;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    TabSheet1: TTabSheet;
    Bevel2: TBevel;
    HLabel6: THLabel;
    CmbChoixTable: THValComboBox;
    HLabel7: THLabel;
    CmbChamps: THValComboBox;
    procedure BChercheClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BImprimerClick(Sender: TObject);
    procedure CmbChoixTableChange(Sender: TObject);
    procedure CmbChampsChange(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    FStChamp    : string; // nom du champ de recherche ( sois _REFINTERNE ou _REFEXTERNE )
    FStNomTable : string; // nom de table de recherche
    FStPrefixe  : string; // prefixe de la table de recherche
    FTOBListe   : TOB;   // TOB contenant le resultat du filtre
    FNomFiltre  : string;
    ObjFiltre   : TObjFiltre; {JP 01/06/07}

    procedure Cherche;  // contruit la requete de recherche
    procedure RempliGrid; // rempli la grille avec les numero manquant
  public
    { Déclarations publiques }
  end;

procedure CPLanceFiche_VisuTresorerie;


implementation

{$R *.DFM}
uses
  UObjEtats;

procedure CPLanceFiche_VisuTresorerie;
var
 lF  : TFVisuEtebac;
 lPP : THPanel;
begin

 lF  := TFVisuEtebac.Create(Application) ;
 lPP := FindInsidePanel;

  if lPP = nil then
  begin
    try
      lF.ShowModal ;
    finally
      lF.Free ;
    end ;
  end
   else
     begin
      InitInside(lF,lPP);
      lF.Show;
    end;
end;


procedure TFVisuEtebac.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
 FTOBListe := TOB.Create('', nil , -1 );
 FNomFiltre := 'RLB_VISU';
  {JP 01/06/07 : Gestion des filtres}
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;

  ObjFiltre := TObjFiltre.Create(Composants, FNomFiltre);
end;

procedure TFVisuEtebac.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if assigned(FTOBListe) then FTOBListe.Free;
end;

procedure TFVisuEtebac.FormResize(Sender: TObject);
begin
 HMTrad.ResizeGridColumns(FListe);
end;

procedure TFVisuEtebac.FormShow(Sender: TObject);
begin
 if V_PGI.AGLDesigning then exit ;
 HMTrad.ResizeGridColumns(FListe);
 ResizeOnglets(Pages) ;
 CmbChoixTable.ItemIndex := 0;
 CmbChamps.ItemIndex     := 0;
 CmbChoixTableChange(self);
 CmbChampsChange(self);
 pages.ActivePageIndex   := 0;
 {JP 01/06/07 : Gestion des filtres
  charge les filtres par defaut
 ChargeFiltre(FNomFiltre,FFiltres,Pages);}
 ObjFiltre.Charger;

end;

procedure TFVisuEtebac.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if Not PFiltres.Enabled then exit ;
 if Key=vk_applique then BEGIN Key:=0 ; BCherche.Click ; if FListe.CanFocus then FListe.SetFocus ; END ;
 if Key=vk_valide then BEGIN Key:=0 ; BCherche.Click ; if FListe.CanFocus then FListe.SetFocus ; END ;
end;

procedure TFVisuEtebac.BChercheClick(Sender: TObject);
begin
 Cherche;
end;

procedure TFVisuEtebac.BAgrandirClick(Sender: TObject);
begin
 Pages.Visible     := not Pages.Visible;
 BAgrandir.Visible := false;
 BReduire.Visible  := true;
end;

procedure TFVisuEtebac.BReduireClick(Sender: TObject);
begin
 Pages.Visible     := not Pages.Visible;
 BAgrandir.Visible := true;
 BReduire.Visible  := false;
end;

procedure TFVisuEtebac.BImprimerClick(Sender: TObject);
begin
 {$IFDEF EAGLCLIENT}
 TObjEtats.GenereEtatGrille (FListe, Caption, False);
 {$ELSE}
 PrintDBGrid( FListe , Nil , 'Contrôle','') ;
 {$ENDIF EAGLCLIENT}
end;

procedure TFVisuEtebac.CmbChoixTableChange(Sender: TObject);
begin
 FStNomTable := CmbChoixTable.Value;
 FStPrefixe  := TableToPrefixe(FStNomTable);
 FStChamp    := FStPrefixe + CmbChamps.Value;
end;

procedure TFVisuEtebac.CmbChampsChange(Sender: TObject);
begin
 CmbChoixTableChange(nil);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/05/2002
Modifié le ... : 06/05/2002
Description .. : Contruction de la requete de recherche
Mots clefs ... :
*****************************************************************}
procedure TFVisuEtebac.Cherche;
var
 lStSQL          : string;
 Q               : TQuery;
 lTOBLigneResult : TOB;
begin


 if ( FStChamp = '' ) or ( FStNomTable = '' ) then exit;

 lStSQL := ' select distinct '              +
            FStChamp                        +
           ' from ' + FStNomTable;


 lStSQL := lStSQL + ' where ( ' + FStPrefixe + '_DATECOMPTABLE >= "'
 + UsDate(EdtMulDateDu) + '" and ' + FStPrefixe + '_DATECOMPTABLE <= "' + UsDate(EdtMulDateAu) + '") ' ;

 if EdtJournal.Text <> '' then
  lStSQL := lStSQL + ' and ' + FStPrefixe + '_JOURNAL = "' + EdtJournal.Text + '"';

 Q          := nil;
 FTOBListe.ClearDetail;

 // suppression des eventuels enregsitrement de la grille
 FListe.BeginUpdate;
 // GCO - 31/05/2006
 FListe.RowCount := 2;
 FListe.Row := 1;
 FListe.VidePile(false);
 FListe.EndUpdate;

 try

   Q := openSQL( lStSQL, true );

   if Q.EOF then exit;

   // remplissage de la TOB avec la valeur de la requete.
   // les valeurs non numerique sont supprime et les chaines sont convertis en nombre pour etre trie
   while not Q.EOF do
    begin
     if IsNumeric(Q.FindField(FStChamp).asString) and (Valeur(Q.FindField(FStChamp).asString) >= Valeur(EdtReference.Text)) then
      begin
        lTOBLigneResult := TOB.Create('',FTOBListe,-1);
        lTOBLigneResult.AddChampSupValeur(FStChamp,Valeur(Q.FindField(FStChamp).asString),false);
      end; // if
       Q.Next;
    end; // while

   FTOBListe.Detail.Sort(FStChamp);

 finally
   if assigned(Q) then Ferme(Q);
 end; // try

 RempliGrid;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/05/2002
Modifié le ... :   /  /
Description .. : recherche des trous dans les numeros
Mots clefs ... :
*****************************************************************}
procedure TFVisuEtebac.RempliGrid;
var
 lTOBResult        : TOB;
 lTOBLigneResult   : TOB;
 lRdNum            : double;
 lRdPrecNum        : double;
 lRdDebNum         : double;
 i                 : integer;
 j                 : integer;
 lRdAmplitude      : double;
begin

 FListe.BeginUpdate;

 try

 if FTOBListe.Detail.Count < 2 then
  begin
   FListe.RowCOunt := 2;
   FListe.Row := 1;
   FListe.VidePile(false);
   exit;
  end;

 lTOBResult := TOB.Create('',nil,-1);

 lRdPrecNum := FTOBListe.Detail[0].GetValue(FStChamp);
 lRdDebNum  := lRdPrecNum;
 lRdNum     := lRdPrecNum;

 if EdtAmplitude.Text <> '' then
  lRdAmplitude := StrToFloat(EdtAmplitude.Text)
   else
    lRdAmplitude := -1;

 for j := 1 to FTOBListe.Detail.Count do
  begin

    if ( lRdNum <> lRdPrecNum ) and (  (lRdNum-1) <> lRdPrecNum  ) then
     begin
      for i := Trunc(lRdNum) to Trunc(lRdPrecNum)-1 do
       begin
        lTOBLigneResult := TOB.Create('',lTOBResult,-1);
        lTOBLigneResult.AddChampSupValeur(FStChamp,FloatToStr(lRdNum),false);
        lRdNum          := lRdNum + 1;
        if ( lRdAmplitude <> -1 ) and ( lRdPrecNum > (lRdDebNum + lRdAmplitude ) ) then break;
       end; // for
     end; // if

   if j = FTOBListe.Detail.Count then break;
   lRdPrecNum := FTOBListe.Detail[j].GetValue(FStChamp);
   lRdNum     := lRdNum + 1;

   if ( lRdAmplitude <> -1 ) and ( lRdPrecNum > (lRdDebNum + lRdAmplitude ) ) then break;

  end; // for

  // ne sert plus en memoire
  FTOBListe.ClearDetail;
  FListe.RowCount := 2;
  FListe.Row := 1;
  FListe.VidePile(false);
  if lTOBResult.Detail.Count <> 0 then
    FListe.RowCount := lTOBResult.Detail.Count+1;

  for i := 0 to lTOBResult.Detail.Count - 1 do
   begin
    FListe.Cells[0,i+1] := lTOBResult.Detail[i].GetValue(FStChamp);
   end;

 lTOBResult.Free;

 finally
  FListe.EndUpdate;
 end;

end;

procedure TFVisuEtebac.BAnnulerClick(Sender: TObject);
begin
  Close;
  if IsInside(Self) then CloseInsidePanel(Self) ;
end;

end.
