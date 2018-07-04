{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 13/05/2004
Modifié le ... :   /  /
Description .. : - FQ 13237 - CA  - 13/05/2004 - Ne prendre que les Immo fi de l'année
Suite..........: - FQ 10319 - PGR - 06/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
Suite..........: -          - MVG - 13/04/2006 - pour compilation de SERIE1
Suite..........: - FQ 17498 - MBO - 13/04/2006 - remplacer la touche F10 par F9 pour exécuter
Suite..........: - FQ 16750 - MBO - 14/04/2006 - voir les comptes mvtés en cpta même si non mvtés immo
Suite..........:            - MVG - 12/07/2006 - pour compilation de SERIE1
Suite..........: - FQ 18815  BTY Colonne Ecart mangée à droite
Suite.......... :- FQ 19903 - MVG 21/05/07 pour que la croix ferme la fenêtre
Mots clefs ... :
*****************************************************************}
unit ImRapCpt;

interface

uses
   SysUtils
   , Classes
   , Graphics
   , Controls
   , Forms
   , Dialogs
   , Grids
   , Hctrls
   , HEnt1
   , ImEnt
   {$IFNDEF SERIE1} // 13/04/2006
{$IFNDEF CMPGIS35}
   , uTofConsEcr
{$ENDIF}
   {$ENDIF}
   , windows  //MVG 13/04/2006
   , ImGenEcr
   , HPanel
   , Uiutil
   , HSysMenu
   , HTB97
   , HStatus
   , UTob
   , ed_tools
   , HMsgBox
   , Mask
   , StdCtrls
   , ExtCtrls;

const COL_COMPTE=0;
      COL_LIBELLE=1;
      COL_SOLDEIMMO=2;
      COL_SOLDECOMPTA=3;
      COL_ECART=4;
type
  TFRapCpt = class(TForm)
    PBouton: TToolWindow97;
    FListe: THGrid;
    FindDialog: TFindDialog;
    Dock: TDock97;
    HMTrad: THSystemMenu;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BZoomDetail: TToolbarButton97;
    PDATERAPPRO: THPanel;
    HLabel1: THLabel;
    DATERAPPROCHEMENT: THCritMaskEdit;
    BCALCUL: TToolbarButton97;
    Type_ecr: TLabel;
    Ecr_normales: TCheckBox;
    Ecr_Simulation: TCheckBox;
    Ecr_Situation: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BCALCULClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    //PGR - 09/02/2006 - Gestion F5 sur liste des écritures
    procedure OnKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BFermeClick(Sender: TObject);
  private
    { Déclarations privées}
    fListeEcriture : TList;
    FindFirst : boolean;
    procedure InitGrille;
    procedure RecupDepuisImmo;
    procedure RecupDepuisEcri;
    procedure AfficheListeEcriture;
    //procedure MajListeAvecCompta;
    procedure InsOrUpdLaListe(fListeE: TList;Compte,Libelle: string; ImDebit,ImCredit,EcDebit,EcCredit: double;Sens: string) ;
    procedure DessineCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
  public
    { Déclarations publiques}
  end;

  TLigneEcritureRapCpt=class(TLigneEcriture)
    Debit2,
    Credit2 : double ;
    Sens : string ;
  end ;


procedure EtatRapprochementCompta;
//function RechercheCompteDansListe(L : TList;ARecordEch : TLigneEcriture) : TLigneEcriture;

implementation

uses
  {$IFDEF eAGLCLient}
  UtileAGL
  {$ELSE}
  PrintDBG
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  {$ENDIF eAGLCLient}

  {$IFDEF SERIE1}
  , S1Util
  , AGLinit
//  , eDispatchS1
    {$IFDEF eAGLCLient}
  , MaineAGL
//  , LanceurFiches
    {$ELSE}
  , FE_Main
    {$ENDIF eAGLCLient}
  {$ELSE}
  , ZEcriMvt_TOF
  , CalcOLE
  {$ENDIF}
  , AMLISTE_TOF
  ;

{$R *.DFM}

procedure EtatRapprochementCompta;
var FRapCpt: TFRapCpt;
    PP : THPanel ;
begin
  FRapCpt := TFRapCpt.Create (Application);
  PP:=FindInsidePanel ;
  if PP=Nil then
  begin
    try
      FRapCpt.ShowModal;
    finally
      FRapCpt.Free;
    end;
  end
  else
  begin
    InitInside(FRapCpt,PP) ;
    FRapCpt.Show ;
  end ;
end;

function CompareCompte (Item1,Item2:Pointer) : integer;
var Ecr1,Ecr2 : TLigneEcritureRapCpt;
begin
  Ecr1 := Item1;Ecr2 := Item2;
  if Ecr1.Compte > Ecr2.Compte then Result := 1
  else if Ecr1.Compte < Ecr2.Compte then Result := -1
  else Result := 0;
end;

(*function RechercheCompteDansListe(L : TList;ARecordEch : TLigneEcriture) : TLigneEcriture;
var ARecord : TLigneEcriture;
    i : integer;
begin
  Result := nil;
  for i := 1 to L.Count do
  begin
    ARecord := L.Items[i-1];
    if (ARecord.Compte=ARecordEch.Compte) then
    begin
      Result := ARecord; break;
    end;
  end;
end;*)

procedure TFRapCpt.FormShow(Sender: TObject);
begin
  DATERAPPROCHEMENT.Text := DateToStr(VHImmo^.Encours.Fin);
  fListeEcriture := TList.Create;
  InitGrille;
  //FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  ecr_Normales.State := cbChecked;
{$IFDEF SERIE1}
  RecupDepuisImmo ;
  RecupDepuisEcri ;
  AfficheListeEcriture;
  BZoomDetail.Enabled := FListe.Cells[0,Fliste.Row]<>'';
  //FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  //PDATERAPPRO.Visible := False;
  type_Ecr.Visible := False;
  ecr_Normales.Visible := False;
  ecr_Simulation.Visible := False;
  ecr_Situation.Visible := False;
{$ELSE}
{$ENDIF}
end;

procedure TFRapCpt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  VideListe(fListeEcriture) ;
  fListeEcriture.Free ;
//  FListe.VidePile(True) ;
  if isInside(Self) then Action:=caFree ;
end;

procedure TFRapCpt.InitGrille;
begin
  Canvas.Brush.Color               :=clWindow;
  fListe.GetCellCanvas             :=DessineCell;
  fListe.ColAligns[COL_COMPTE]     :=taCenter;
  fListe.ColAligns[COL_LIBELLE]    :=taLeftJustify;
  fListe.ColAligns[COL_SOLDEIMMO]  :=taRightJustify;
  fListe.ColAligns[COL_SOLDECOMPTA]:=taRightJustify;
  fListe.ColAligns[COL_ECART]      :=taRightJustify;
end;

procedure TFRapCpt.RecupDepuisImmo;
var ParamE : TParamEcr; ARecord : TLigneEcriture; i : integer; ListeEcheances : TList; Q: TQuery ;
  DateRappro : TDateTime;
begin
{$IFDEF SERIE}
  DateRappro := StrToDate(DATERAPPROCHEMENT.Text);
//  DateRappro := VHImmo^.Encours.Fin;
{$ELSE}
  DateRappro := StrToDate(DATERAPPROCHEMENT.Text);
{$ENDIF}
  //Récupération des dotations
  ParamE := TParamEcr.Create;
  ListeEcheances := TList.Create;
  try
    ParamE.Journal   :='';
    ParamE.Libelle   :='';
//    ParamE.Date      :=VHImmo^.Encours.Fin;
    ParamE.Date      := DateRappro;
    ParamE.DateCalcul:= DateRappro;
    ParamE.bDotation :=True;
    SourisSablier;
    InitMove (1,'');
    CalculEcrituresDotation(ListeEcheances,'',ParamE,4);
    FiniMove;
    for i:=0 to ListeEcheances.Count - 1 do
    begin
      ARecord:=ListeEcheances.Items[i] ;
      if (ARecord.Compte='') or (ARecord.Auxi<>'') then continue;
      InsOrUpdLaListe(fListeEcriture,ARecord.Compte,ARecord.Libelle,ARecord.Debit,ARecord.Credit,0,0,'') ;
    end ;
  finally
    VideListeEcritures(ListeEcheances) ;
    ListeEcheances.Free ;
    ParamE.Free;
  end ;

// CA - 04/09/2002 - Ces écritures ne doivent pas être prises en comptes
(*  //récupération des échéances
  ParamE := TParamEcr.Create;
  ListeEcheances := TList.Create;
  try
    ParamE.Journal:='OD' ; //aucune importance simulation de génération d'écriture
    InitMove (1,'');
    CalculEcrituresEcheances(ListeEcheances,'',ParamE,nil ,0,VHImmo^.Encours.deb,VHImmo^.Encours.fin,true);
    FiniMove;
    for i:=0 to ListeEcheances.Count - 1 do
    begin
      ARecord:=ListeEcheances.Items[i] ;
      if (ARecord.Compte='') or (ARecord.Auxi<>'') then continue;
      InsOrUpdLaListe(fListeEcriture,ARecord.Compte,ARecord.Libelle,ARecord.Debit,ARecord.Credit,0,0,'') ;
    end ;
    SourisNormale;
  finally
    VideListeEcritures(ListeEcheances) ;
    ListeEcheances.Free ;
    ParamE.Free;
  end ;*)

  //récupération des immo fi  achats
  { CA - 13/05/2004 - FQ 13237 - Ne prendre que les Immo fi. de l'année et antérieures }
  Q:=OpenSql('SELECT * FROM IMMO WHERE I_NATUREIMMO="FI" AND I_DATEPIECEA<="'+USDateTime(VHImmo^.Encours.Fin)+'" AND I_ETAT<>"FER"',false) ;
  while not Q.Eof do
  begin
    InsOrUpdLaListe(fListeEcriture,Q.FindField('I_COMPTEIMMO').AsString ,'',Q.FindField('I_MONTANTHT').AsFloat ,0,0,0,'') ;
    Q.Next ;
  end;
  ferme(Q) ;
end;

procedure TFRapCpt.RecupDepuisEcri;
var Q : TQuery; LaTob,T1: Tob ; i: integer ; StChamps,StWhere,wSens: string ; MttDeb,MttCre: double ;

  {$IFDEF SERIE1}
    // FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
    DateRappro : TDateTime;
  {$ELSE}
    TResult : TabloExt;
    DateRappro : TDateTime;
    typeecr : string; // MVG 12/07/2006
  {$ENDIF}
begin
  LaTob:=Tob.Create('GENERAUX',nil,-1) ;
  {$IFDEF SERIE1}
  // FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  DateRappro := StrToDate(DATERAPPROCHEMENT.Text);
  StChamps:='G_GENERAL,G_LIBELLE' ; //,G_TOTALDEBIT,G_TOTALCREDIT ' ;
  {$ELSE}
  DateRappro := StrToDate(DATERAPPROCHEMENT.Text);
  StChamps:='G_GENERAL,G_LIBELLE,G_TOTDEBE,G_TOTCREE,G_SENS ' ;
  {$ENDIF}
  try
    //Tous les compte de nature IMO
    Q:=OpenSQL ('SELECT '+StChamps+' FROM GENERAUX WHERE G_NATUREGENE="IMO"', TRUE);
    LaTob.LoadDetailDB('GENERAUX','','',Q,false,true) ;
    ferme(Q) ;
    //Tous les comptes utilisé pour les immo hors ceux de nature IMO
    StWhere:='' ;
    for i:=0 to fListeEcriture.Count-1 do
    begin
      if StWhere<>'' then StWhere:=StWhere+' OR ' ;
      StWhere:=StWhere+' G_GENERAL="'+TLigneEcritureRapCpt(fListeEcriture.Items[i]).Compte+'"' ;
    end;
    if StWhere<>'' then
    begin
      Q:=OpenSQL ('SELECT '+StChamps+' FROM GENERAUX WHERE G_NATUREGENE<>"IMO" AND ( '+StWhere+' )', TRUE);
      LaTob.LoadDetailDB('GENERAUX','','',Q,true,true) ;
      ferme(Q) ;
    end ;
    // Verification de la TOB
    for i:=0 to LaTob.detail.Count-1 do
    begin
      T1:=LaTob.Detail[i] ;
      {$IFDEF SERIE1}
      wSens :='C' ;
      // FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
      //MttCre:=S1_GetCumul('E','C','<='+FloatToStr(VHImmo^.EnCours.Fin),'','',T1.GetValue('G_GENERAL'),'','','','') ;  //T1.GetValue('G_TOTALCREDIT') ;
      //MttDeb:=S1_GetCumul('E','D','<='+FloatToStr(VHImmo^.EnCours.Fin),'','',T1.GetValue('G_GENERAL'),'','','','') ; //T1.GetValue('G_TOTALDEBIT') ;
      MttCre:=S1_GetCumul('E','C','<='+FloatToStr(DateRappro),'','',T1.GetValue('G_GENERAL'),'','','','') ;  //T1.GetValue('G_TOTALCREDIT') ;
      MttDeb:=S1_GetCumul('E','D','<='+FloatToStr(DateRappro),'','',T1.GetValue('G_GENERAL'),'','','','') ; //T1.GetValue('G_TOTALDEBIT') ;
      {$ELSE}
      wSens :=T1.GetValue('G_SENS') ;
      // FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
     //GetCumul('GEN',T1.GetValue('G_GENERAL'),'','N','','',VHImmo^.Encours.Code,
     //  VHImmo^.Encours.Deb,DateRappro,False,False,nil,TResult,False);
     if ecr_Normales.State = cbChecked then
        typeecr := 'N';
      if ecr_Situation.State = cbChecked then
        typeecr := typeecr + 'U';
      if ecr_Simulation.State = cbChecked then
        typeecr := typeecr + 'S';
      GetCumul('GEN',T1.GetValue('G_GENERAL'),'',typeecr,'','',VHImmo^.Encours.Code,
        VHImmo^.Encours.Deb,DateRappro,False,False,nil,TResult,False);
//      MttCre := T1.GetValue('G_TOTCREE') ;
//      MttDeb := T1.GetValue('G_TOTDEBE');
      MttCre := TResult[2];
      MttDeb := TResult[1];
      {$ENDIF}
      InsOrUpdLaListe(fListeEcriture,T1.GetValue('G_GENERAL'),T1.GetValue('G_LIBELLE'),0,0,MttDeb,MttCre,wSens) ;
    end;
  finally
    LaTob.Free ;
  end ;
end;


procedure TFRapCpt.InsOrUpdLaListe(fListeE: TList;Compte,Libelle: string; ImDebit,ImCredit,EcDebit,EcCredit: double;Sens: string) ;
var i: integer ; bTrouve: boolean ; ARecord: TLigneEcritureRapCpt ;
begin
  bTrouve := false; i:=0 ;
  while (i<=fListeE.Count-1) and not bTrouve do
  begin
    bTrouve:=(TLigneEcritureRapCpt(fListeE.Items[i]).Compte=Compte) ;
    if not bTrouve then inc(i);
  end;
  if bTrouve then
  begin
    ARecord:=TLigneEcritureRapCpt(fListeE.Items[i]) ;
    ARecord.Debit  :=ARecord.Debit+ImDebit;
    ARecord.Credit :=ARecord.Credit+ImCredit;
    ARecord.Debit2 :=ARecord.Debit2+EcDebit;
    ARecord.Credit2:=ARecord.Credit2+EcCredit;
    if ARecord.Libelle='' then ARecord.Libelle:=Libelle ;
    if Sens<>''           then ARecord.Sens:=Sens ;
  end
  else
  begin
    // fq 16750if (ImDebit+ImCredit+EcDebit+EcCredit<>0) then
    // fq 16750 begin

      ARecord        :=TLigneEcritureRapCpt.Create ;
      ARecord.Compte :=Compte ;
      ARecord.Libelle:=Libelle ;
      ARecord.Debit  :=ImDebit ;
      ARecord.Credit :=ImCredit ;
      ARecord.Debit2 :=EcDebit;
      ARecord.Credit2:=EcCredit;
      if Sens<>'' then ARecord.Sens:=Sens ;
      fListeEcriture.Add(ARecord);
    // fq  16750 end ;
  end;
end ;


procedure TFRapCpt.AfficheListeEcriture;
var i : integer; ARecord : TLigneEcritureRapCpt; SoldeImmo, SoldeCpta, SoldeDiff : double; j:integer;

begin
  j:=0;
  // fq 16750 fListe.RowCount:=2 ;
  fListe.RowCount:=1 ;
  // fq 16750 if fListeEcriture.Count<>0 then fListe.RowCount:=fListeEcriture.Count+1;
  fListeEcriture.Sort(CompareCompte);
  for i:=1 to fListeEcriture.Count do
  begin
    ARecord:=fListeEcriture.Items[i-1];

    // fq 16750 déplacement des lignes
    SoldeCpta:=ARecord.Debit2-ARecord.Credit2 ;
    SoldeImmo:=ARecord.Debit-ARecord.Credit ;

    if (SoldeCpta <> 0) or (SoldeImmo <> 0) then
    begin
    inc(j);
    fListe.RowCount:=fListe.RowCount+1;

    fListe.Cells[COL_COMPTE,j] :=ARecord.Compte;
    fListe.Cells[COL_LIBELLE,j]:=ARecord.Libelle;
    (*if ARecord.Sens<>'C' then
    begin*)
    //fq 16750   SoldeCpta:=ARecord.Debit2-ARecord.Credit2 ;
    //fq 16750  SoldeImmo:=ARecord.Debit-ARecord.Credit ;
    (*end
    else
    begin
      SoldeCpta:=ARecord.Credit2-ARecord.Debit2 ;
      SoldeImmo:=ARecord.Credit-ARecord.Debit ;
    end ;*)
    SoldeDiff:=SoldeCpta-SoldeImmo ;
    fListe.Cells[COL_SOLDEIMMO,j]  :=StrfMontant(SoldeImmo,13,V_PGI.OkDecV,'',True);
    fListe.Cells[COL_SOLDECOMPTA,j]:=StrfMontant(SoldeCpta,13,V_PGI.OkDecV,'',True);
    fListe.Cells[COL_ECART,j]      :=StrfMontant(SoldeDiff,13,V_PGI.OkDecV,'',True);
    end;
  end;

  // FQ 18815 Provoquer le réaffichage
  HMTrad.ResizeDBGrid := True;
  if ((HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid) and (V_PGI.Outlook)) then
      HMTrad.ResizeGridColumns (fListe);
end;

procedure TFRapCpt.DessineCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
  if (fListeEcriture = nil) or (fListeEcriture.Count<=0) then exit;
  if (fListeEcriture.Count > 0) and (ARow > 0)  then
  begin
    if (Valeur (fListe.Cells[COL_ECART,ARow])=0)
    then
    else
      if (gdSelected in AState) then Canvas.Brush.Color := clHighLight
      else Canvas.Brush.Color := AltColors[V_PGI.NumAltCol];
  end;
end ;

procedure TFRapCpt.FListeDblClick(Sender: TObject);
var Compte: string;
  {$IFDEF SERIE1}
  Critere: string ;
  {$ELSE}
  {$ENDIF}
begin
  Compte := FListe.Cells[0,Fliste.Row];
  if Compte = '' then exit;
  {$IFDEF SERIE1}
  //XMG 30/04/03 début
  // FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  //Critere:='E_DATECOMPTABLE='+DateToStr(VHImmo^.EnCours.Deb)+';E_DATECOMPTABLE_='+DateToStr(VHImmo^.EnCours.Fin)+';'
  //        +'E_GENERAL='+Compte+';E_GENERAL_='+Compte+';|' ;
  Critere:='E_DATECOMPTABLE='+DateToStr(VHImmo^.EnCours.Deb)+';E_DATECOMPTABLE_='+ DATERAPPROCHEMENT.Text+';'
          +'E_GENERAL='+Compte+';E_GENERAL_='+Compte+';|' ;
  //ConsultEcriture_Show(Nil, false,Critere ,'', taConsult);
  AGLLanceFiche('C','MULECR',Critere,'',Actiontostring(taConsult)+';') ;
  //XMG 30/04/03 fin
  //MulTG_Show(nil,'GENERAUX',Critere) ;
  {$ELSE}
  // FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  //ZoomEcritureMvt(Compte, ImGeneTofb,'MULMMVTS' ) ;
{$IFNDEF CMPGIS35}
  OperationsSurComptes(Compte,VHIMMO^.Encours.code, '', '',false) ;
{$ENDIF}
  {$ENDIF}
end;

procedure TFRapCpt.BImprimerClick(Sender: TObject);
(*var
   titre: string ;  *)
begin
//  Titre:=caption ;

  {$IFDEF eAGLCLient}
  PrintDBGrid(caption,FListe.name,'') ;
 {$ELSE}
  //FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  //PrintDBGrid(FListe,Nil,caption,PDATERAPPRO,'') ;
  PrintDBGrid(FListe,PDATERAPPRO,caption,'') ;
  {$ENDIF eAGLCLient}

end;

procedure TFRapCpt.BRechercherClick(Sender: TObject);
begin FindFirst:=True; FindDialog.Execute ; end;

procedure TFRapCpt.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FindFirst); end;

procedure TFRapCpt.FListeEnter(Sender: TObject);
begin BZoomDetail.Enabled := FListe.Cells[0,Fliste.Row]<>''; end;

procedure TFRapCpt.BAideClick(Sender: TObject);
begin CallHelpTopic(Self); end;

procedure TFRapCpt.FormCreate(Sender: TObject);
begin
{$IFDEF SERIE1}
  //FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  //Caption:='Rapprochement avec la comptabilité au '+DateToStr(VHImmo^.EnCours.Fin) ;
HelpContext:=541500 ;
{$ELSE}
HelpContext:=2415000 ;
{$ENDIF}

end;

procedure TFRapCpt.BCALCULClick(Sender: TObject);
begin
  // Contrôle de la validité de la date de rapprochement
  if ((not IsValidDate(DATERAPPROCHEMENT.Text)) or
        (StrToDate(DATERAPPROCHEMENT.Text)>VHImmo^.Encours.Fin) or
        (StrToDate(DATERAPPROCHEMENT.Text)<VHImmo^.Encours.Deb)) then
    PGIBox ('Date de rapprochement non valide.#10#13Veuillez saisir une date de l''exercice en cours',Caption)
  //FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
  else if ((ecr_Normales.State = cbUnChecked) and (ecr_Simulation.State = cbUnChecked) and
    (ecr_Situation.State = cbUnChecked)) then
    PGIBox ('Veuillez choisir au moins un type d''écritures.',Caption)
  else
  begin
    VideListe(fListeEcriture) ;
    RecupDepuisImmo ;
    RecupDepuisEcri ;
    AfficheListeEcriture;
    BZoomDetail.Enabled := FListe.Cells[0,Fliste.Row]<>'';
  end;
end;

procedure TFRapCpt.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//FQ 10319 - PGR 6/01/2006 - Rapprochement encours d'exercice avec écrit. de simulation
//{$IFNDEF SERIE1}
  // fq 17498 - mbo - 13.04.2006
  //if Key=VK_F10 then BCALCULClick(nil);
  if Key=VK_F9 then BCALCULClick(nil);
//{$ENDIF}
end;
//PGR - 09/02/2006 - Gestion F5 sur liste des écritures
procedure TFRapCpt.OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);

begin

  if (Key=VK_F5) then
  begin;
    FlisteDblClick(nil);
    Key := 0;
  end;

end;

procedure TFRapCpt.BFermeClick(Sender: TObject);
begin
// FQ 19903 - MVG 21/05/07
  Close ;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;

end;

end.
