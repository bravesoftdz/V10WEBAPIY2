{***********UNITE*************************************************
Auteur  ...... : C.B
Créé le ...... : 19/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPLANCHARGE ()
Mots clefs ... : TOF;AFPLANCHARGE
*****************************************************************}
Unit UTOFAFPLANCHARGE;

Interface

Uses StdCtrls,
     Controls,
     Classes,

{$IFDEF EAGLCLIENT}
     {UtileAGL, QRGrid,}
     MaineAGL, UtileAGL,
{$ELSE}
     db, dbtables, PrintDBG, FE_Main, EdtREtat,
{$ENDIF}
     CalcOleGenericAff,
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
     UTOF, UTOb, AFPlanningCst, UtilPlanning, dialogs,
     paramsoc, Graphics, Grids, windows, HPanel, HTB97,
     DicoAF, AffaireUtil, UtilTaches, ed_tools, Vierge,
     UtilPlanDeCharge, uafo_ressource, utilRessource;


Type

  TOF_AFPLANCHARGE_ANCETRE = Class (TOF)
    procedure OnClose;  override;
    procedure InitTof;

    private

      fGD_PlanCharge  : THGrid;
      fGD_TOTAUX      : THGrid;
      fTobPlanCharge  : Tob;
      fStAffaire      : String;
      fStLibAffaire   : String;
      fBlGestionDec   : Boolean;
      fPEntete        : THPanel;
      fBVerif         : TToolbarButton97;
      fBRefresh       : TToolbarButton97;
      fBValider       : TToolbarButton97;
      fBImprimer      : TToolbarButton97;

      procedure FormatPCGrid; virtual;
      procedure FormatTob(var pTob : Tob);

      // Chargement
      procedure LoadGrid;
      procedure LoadGeneral;
      procedure LoadTaches; virtual; abstract;
      procedure LoadTotal;
      procedure RechercheRes(pTobTache : Tob);
      procedure AddRessource(pTobMere : Tob; pStRes, pStNumTache, pStFonction, pStTiers, pStAff0, pStAff1, pStAff2, pStAff3, pStAvenant : String);

      // controle de saisie
      procedure fGD_PlanChargeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      Procedure GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
      Procedure GetCellCanvasTotaux( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
      procedure fGD_PlanChargeSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
      procedure BVerifClick(Sender: TObject);
      procedure BRefreshClick(Sender: TObject);
      procedure BValiderClick(Sender: TObject);
      procedure BImprimerClick(Sender: TObject);

      function Verif(pBoAsk : Boolean) : Boolean;

      // form
      procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

      // sauvegarde
      procedure SauvePlanDeCharge;
      function PrepareSauvTob(pTobTaches, pTobLignesPla, pTobLigne : Tob; var pInNumLigne : Integer; TQR : TQRProgressForm) : Boolean; virtual; abstract;

      procedure CalculTotal(aRow : Integer);
      procedure CalculEcart(aRow : Integer);
  end;

  TOF_AFPLANCHARGE_QTE = Class (TOF_AFPLANCHARGE_ANCETRE)
    private
      procedure FormatPCGrid; override;
      procedure LoadTaches; override;
      procedure LoadFonction(pTobTache : Tob);
      procedure LoadRessources(pTobTache : Tob);
      function PrepareSauvTob(pTobTaches, pTobLignesPla, pTobLigne : Tob; var pInNumLigne : Integer; TQR : TQRProgressForm) : Boolean; override;
  end;

  TOF_AFPLANCHARGE_MNT = Class (TOF_AFPLANCHARGE_ANCETRE)
    private
      procedure FormatPCGrid; override;
      procedure LoadTaches; override;
      procedure LoadFonction(pTobTache : Tob);
      function PrepareSauvTob(pTobTaches, pTobLignesPla, pTobLigne : Tob; var pInNumLigne : Integer; TQR : TQRProgressForm) : Boolean; override;
  end;

  TOF_AFPLANCHARGE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private
      fTofPlanCharge  : TOF_AFPLANCHARGE_ANCETRE;
      fStAffaire      : String;
      fStLibAffaire   : String;
  end;
                       
Function AFLanceFicheAFPlanCharge(Lequel, Argument : String) : String;

const
	TexteMsgTache: array[1..5] of string 	= (
          {1}        'Les quantités prévues de %s ne sont pas cohérentes',
          {2}        'Les quantités prévues sont correctes',
          {3}        'Erreur lors de la préparation de la sauvegarde',
          {4}        'Erreur lors de la sauvegarde du plan de charge',
          {5}        'Les quantités prévues de %s ne sont pas cohérentes, voulez vous quand même valider ?');

Implementation

Function AFLanceFicheAFPlanCharge(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANCHARGE','', Lequel, Argument);
end;

{******************************************************************************}
{* TOF_AFPLANCHARGE                                                           *}
{******************************************************************************}
procedure TOF_AFPLANCHARGE.OnLoad;
begin
  Inherited;
  fTofPlanCharge.InitTof;
  fTofPlanCharge.LoadGrid;
end ;

procedure TOF_AFPLANCHARGE.OnArgument (S : String);
var
  vStTmp          : String;
  vStChamp        : String;
  vStValeur       : String;

begin
  Inherited;

  // traitement des arguments
  vStTmp:= (Trim(ReadTokenSt(S)));
  While (vStTmp <> '') do
  Begin
    If (vStTmp <> '') then
    Begin
      DecodeArgument(vStTmp, vStChamp, vStValeur);
      if vStChamp = 'AFFAIRE' then fStAffaire := vStValeur
      else if vStChamp = 'LIBAFFAIRE' then fStLibAffaire := vStValeur
      else if vStChamp = 'PLANCHARGEQTE' then
      begin
        fTofPlanCharge := TOF_AFPLANCHARGE_QTE.Create(ecran, true);
        fTofPlanCharge.ecran := ecran;
        fTofPlanCharge.ecran.Caption := traduitGA('Plan de Charge en Quantité');
        fTofPlanCharge.fStAffaire := fStAffaire;
        fTofPlanCharge.fStLibAffaire := fStLibAffaire;
      end
      else if vStChamp = 'PLANCHARGEMNT' then
      begin
        fTofPlanCharge := TOF_AFPLANCHARGE_MNT.Create(ecran, true);
        fTofPlanCharge.ecran := ecran;
        fTofPlanCharge.ecran.Caption := traduitGA('Plan de Charge en Montant');
        fTofPlanCharge.fStAffaire := fStAffaire;
        fTofPlanCharge.fStLibAffaire := fStLibAffaire;
      end;

    End;
    vStTmp:=(Trim(ReadTokenSt(S)));
  End;
end;
                            
procedure TOF_AFPLANCHARGE.OnClose;
begin
  Inherited;

  if ecran.ModalResult = MrCancel then
    TfVierge(Ecran).retour := 'FALSE'
  else
    TfVierge(Ecran).retour := 'TRUE';

  if assigned(fTofPlanCharge) then
    begin
      fTofPlanCharge.Close;
      fTofPlanCharge.Free;
    end;             
    
end;

{******************************************************************************}
{* TOF_AFPLANCHARGE_ANCETRE                                                   *}
{******************************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.InitTof;
begin

  fGD_PlanCharge := THGrid(getcontrol('GD_PlanCharge'));
  fGD_TOTAUX := THGrid(getControl('GD_TOTAUX'));
  fPEntete := THPanel(getControl('PENTETE'));
  fBVerif := TToolbarButton97(getControl('BVERIF'));
  fBRefresh := TToolbarButton97(getControl('BREFRESH'));
  fBValider := TToolbarButton97(getControl('BVALIDER'));
  fBImprimer := TToolbarButton97(getControl('BIMPRIMER'));

  fGD_PlanCharge.OnCellExit := fGD_PlanChargeCellExit;
  fGD_PlanCharge.OnSelectCell := fGD_PlanChargeSelectCell;
  fGD_PlanCharge.GetCellCanvas := GetCellCanvas;
  fBVerif.OnClick := BVerifClick;
  fBRefresh.OnClick := BRefreshClick;
  fBValider.OnClick := BValiderClick;
  fBImprimer.OnClick := BImprimerClick;

  ecran.OnKeyDown := FormKeyDown;
  fBlGestionDec := getparamsoc('SO_AFPDCDEC');

//{$IFDEF EAGLCLIENT}
//  setcontrolVisible('BIMPRIMER', false);
//{$ELSE}
  setcontrolVisible('BIMPRIMER', true);
//{$ENDIF}
end;

procedure TOF_AFPLANCHARGE_ANCETRE.OnClose;
begin
  Inherited;
  if assigned(fTobPlanCharge) then
    begin
      fTobPlanCharge.Free;
      fTobPlanCharge := nil;
    end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Chargement des données de la grille du plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.LoadGrid;
begin
  LoadGeneral;                              
  LoadTaches;
  LoadTotal;
end;

procedure TOF_AFPLANCHARGE_ANCETRE.LoadGeneral;
var
  vQr               : TQuery;
  S                 : String;
  vStAff0, vStAff1  : String;
  vStAff2, vStAff3  : String;
  vStAvenant        : String;
  vStTiers          : String;

begin
  // affichage affaire et tiers
  CodeAffaireDecoupe(fStAffaire, vStAff0, vStAff1, vStAff2,
                     vStAff3,vStAvenant, taModif, false);

  // C.B prise en compte des affaires modeles
  TeststCleAffaire(fStAffaire,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant,vStTiers, False, True, False, False);

  SetControltext('ATA_AFFAIRE', fStAffaire);
  SetControltext('ATA_AFFAIRE1', vStAff1);
  SetControltext('ATA_AFFAIRE2', vStAff2);
  SetControltext('ATA_AFFAIRE3', vStAff3);
  SetControltext('ATA_AVENANT', vStAvenant);
  SetControltext('ATA_TIERS', vStTiers);

  ChargeCleAffaire( THEdit(GetControl('ATA_AFFAIRE0')),
                    THEdit(GetControl('ATA_AFFAIRE1')),
                    THEdit(GetControl('ATA_AFFAIRE2')),
                    THEdit(GetControl('ATA_AFFAIRE3')),
                    THEdit(GetControl('ATA_AVENANT')),
                    Nil, taConsult, fStAffaire, False);

  S := 'SELECT AFF_LIBELLE, T_LIBELLE FROM TIERS, AFFAIRE ';
  S := S + 'WHERE AFF_AFFAIRE = "' + fStAffaire + '"';
  S := S + 'AND AFF_TIERS = T_TIERS';

  vQr := nil;
  Try
    vQR := OpenSql(S, True);
    if Not vQR.Eof then
      begin
        SetControlText('AFF_LIBELLE', vQR.FindField('AFF_LIBELLE').AsString);
        SetControlText('T_LIBELLE', vQR.FindField('T_LIBELLE').AsString);
      end;
  finally
    if vQR <> nil then ferme(vQR);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/11/2002
Modifié le ... :
Description .. : calcul des totaux du plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.LoadTotal;
var
  i,j : Integer;

begin
  fGD_TOTAUX.Cells[cInColPrevu,0]   := '0';
  fGD_TOTAUX.Cells[cInColAffecte,0] := '0';
  fGD_TOTAUX.Cells[cInColRealise,0] := '0';
  fGD_TOTAUX.Cells[cInColRafTh,0]   := '0';
  fGD_TOTAUX.Cells[cInColRaf,0]     := '0';
  fGD_TOTAUX.Cells[cInColTotal,0]   := '0';
  fGD_TOTAUX.Cells[cInColEcart,0]   := '0';

  fGD_TOTAUX.CellValues[4,0]    := '0';
  fGD_TOTAUX.CellValues[5,0]    := '0';
  fGD_TOTAUX.CellValues[6,0]    := '0';
  fGD_TOTAUX.CellValues[7,0]    := '0';
  fGD_TOTAUX.CellValues[8,0]    := '0';
  fGD_TOTAUX.CellValues[9,0]    := '0';
  fGD_TOTAUX.CellValues[10,0]   := '0';
  fGD_TOTAUX.CellValues[11,0]   := '0';
  fGD_TOTAUX.CellValues[12,0]   := '0';
  fGD_TOTAUX.CellValues[13,0]   := '0';
  fGD_TOTAUX.CellValues[14,0]   := '0';
  fGD_TOTAUX.CellValues[15,0]   := '0';

  fGD_TOTAUX.Cells[0,0]   := 'Totaux';
  fGD_TOTAUX.ColWidths[0]  := fGD_PlanCharge.ColWidths[0];
  for i := 1 to fGD_PlanCharge.colcount - 1 do
  begin
    fGD_TOTAUX.ColWidths[i] := fGD_PlanCharge.ColWidths[i];
    fGD_TOTAUX.ColAligns[i] := TaRightJustify;
    fGD_TOTAUX.ColTypes[i]  := 'R';

    if (fGD_PlanCharge.Cells[cInColLigFct,i] = 'X') then
    begin
      if fGD_PlanCharge.Cells[cInColPrevu,i] <> '' then fGD_TOTAUX.Cells[cInColPrevu,0]   := FormatJour(valeur(fGD_TOTAUX.Cells[cInColPrevu,0]) + valeur(fGD_PlanCharge.Cells[cInColPrevu,i]));
      if fGD_PlanCharge.Cells[cInColAffecte,i] <> '' then fGD_TOTAUX.Cells[cInColAffecte,0] := FormatJour(valeur(fGD_TOTAUX.Cells[cInColAffecte,0]) + valeur(fGD_PlanCharge.Cells[cInColAffecte,i]));
      if fGD_PlanCharge.Cells[cInColRealise,i] <> '' then fGD_TOTAUX.Cells[cInColRealise,0] := FormatJour(valeur(fGD_TOTAUX.Cells[cInColRealise,0]) + valeur(fGD_PlanCharge.Cells[cInColRealise,i]));
      if fGD_PlanCharge.Cells[cInColRafTh,i] <> '' then fGD_TOTAUX.Cells[cInColRafTh,0]   := FormatJour(valeur(fGD_TOTAUX.Cells[cInColRafTh,0]) + valeur(fGD_PlanCharge.Cells[cInColRafTh,i]));
      if fGD_PlanCharge.Cells[cInColRaf,i] <> '' then fGD_TOTAUX.Cells[cInColRaf,0]     := FormatJour(valeur(fGD_TOTAUX.Cells[cInColRaf,0]) + valeur(fGD_PlanCharge.Cells[cInColRaf,i]));
      if fGD_PlanCharge.Cells[cInColTotal,i] <> '' then fGD_TOTAUX.Cells[cInColTotal,0]   := FormatJour(valeur(fGD_TOTAUX.Cells[cInColTotal,0]) + valeur(fGD_PlanCharge.Cells[cInColTotal,i]));
      if fGD_PlanCharge.Cells[cInColEcart,i] <> '' then fGD_TOTAUX.Cells[cInColEcart,0]   := FormatJour(valeur(fGD_TOTAUX.Cells[cInColEcart,0]) + valeur(fGD_PlanCharge.Cells[cInColEcart,i]));

      for j := 4 to 15 do
        if ((fGD_PlanCharge.CellValues[j,i] <> '') and (fGD_PlanCharge.CellValues[j,i] <> #0)) then
          fGD_TOTAUX.Cells[j,0] := FormatJour(valeur(fGD_TOTAUX.Cells[j,0]) + valeur(fGD_PlanCharge.Cells[j,i]));
    end;
  end;

  // suppression des 0
  for i := 4 to 15 do
    if valeur(fGD_TOTAUX.Cells[i,0]) = 0 then
      fGD_TOTAUX.Cells[i,0] := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 12/11/2002
Modifié le ... :
Description .. : Recherche des ressources pour lesquelles ont a saisi
                 de l'activité et qui ne sont pas saisies dans
                 la tâche
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.RechercheRes(pTobTache : Tob);
var
  S           : String;
  vQr         : TQuery;
  vTob        : Tob;
  i           : Integer;
  vQrRes      : TQuery;
  vTobRes     : Tob;
  vTobMere    : Tob;
  vStAff0     : String;
  vStAff1     : String;
  vStAff2     : String;
  vStAff3     : String;
  vStAvenant  : String;

begin
  S := 'SELECT ACT_RESSOURCE ';
  S := S + ' FROM ACTIVITE, ARTICLE ';
  S := S + ' WHERE ACT_AFFAIRE = "' + fStAffaire + '"';

  if GetParamSoc('SO_AFCLIENT') = cInClientAlgoe then
    S := S + ' AND GA_FAMILLENIV1 <> "PRO"';

  S := S + ' AND GA_ARTICLE = ACT_ARTICLE';
  S := S + ' AND GA_ARTICLE = "' + pTobTache.GetValue('ATA_ARTICLE') + '"';
  S := S + ' AND ACT_ACTIVITEREPRIS <> "A"';
  S := S + ' AND ACT_ETATVISA = "VIS" ';
  S := S + ' GROUP BY ACT_RESSOURCE';

  vQr := nil;
  vTob := TOB.create('Les Ressources', nil, -1);
                            
  Try
    vQr := OpenSql(S,True);
    if Not vQr.Eof then
    begin

      // chargement des ressources existantes
      S := 'SELECT ATR_AFFAIRE, ATR_NUMEROTACHE, ATR_RESSOURCE ';
      S := S + ' FROM TACHERESSOURCE ';
      S := S + ' WHERE ATR_AFFAIRE = "' + fStAffaire + '"';
      S := S + ' AND ATR_NUMEROTACHE = ' + floatToStr(pTobTache.GetValue('ATA_NUMEROTACHE'));

      vQrRes := nil;
      vTobRes := TOB.create('Mes Ressources', nil, -1);
      Try
        vQrRes := OpenSql(S,True);
        if Not vQrRes.Eof then
          vTobRes.LoadDetailDB('Mes Ressources','','',vQrRes,False,True);
      finally
        ferme(vQrRes);
      end;

      vTob.LoadDetailDB('Les Ressources','','',vQr,False,True);
      vTobMere := TOB.create('Mere', nil, -1);
      try
        CodeAffaireDecoupe(fStAffaire, vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);

        for i := 0 to vTob.Detail.Count - 1 do
        begin
          // ajout la ressource à la liste des ressources à ajouter
          if vTobRes.FindFirst(['ATR_AFFAIRE', 'ATR_NUMEROTACHE', 'ATR_RESSOURCE'], [fStAffaire, pTobTache.GetValue('ATA_NUMEROTACHE'), vTob.detail[i].GetValue('ACT_RESSOURCE')], true) = nil then
            AddRessource(vTobMere, vTob.detail[i].GetValue('ACT_RESSOURCE'),
                         pTobTache.GetValue('ATA_NUMEROTACHE'),
                         pTobTache.GetValue('ATA_FONCTION'),
                         pTobTache.GetValue('ATA_TIERS'),
                         vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant);
        end;
        if vTobMere.detail.count > 0 then
        begin
          vTobMere.SetAllModifie(true);
          vTobMere.InsertDB(nil);
        end;
      finally
        vTobMere.Free;
      end;
    end;

  Finally
    ferme(vQr);
  End;
end;

procedure TOF_AFPLANCHARGE_ANCETRE.AddRessource(pTobMere : Tob; pStRes,
                                                pStNumTache, pStFonction, pStTiers,
                                                pStAff0, pStAff1, pStAff2, pStAff3, pStAvenant : String);
var
  vTob        : Tob;
begin
  vTob := Tob.Create('TACHERESSOURCE', pTobMere, -1 );
  vTob.PutValue('ATR_NUMEROTACHE', pStNumTache);
  vTob.PutValue('ATR_TIERS', pStTiers);
  vTob.PutValue('ATR_AFFAIRE', fStAffaire);
  vTob.PutValue('ATR_AFFAIRE0', pStAff0);
  vTob.PutValue('ATR_AFFAIRE1', pStAff1);
  vTob.PutValue('ATR_AFFAIRE2', pStAff2);
  vTob.PutValue('ATR_AFFAIRE3', pStAff3);
  vTob.PutValue('ATR_AVENANT', pStAvenant);
  vTob.PutValue('ATR_RESSOURCE', pStRes);
  vTob.PutValue('ATR_FONCTION', pStFonction);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : creation des champs de la tob virtuelle
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.FormatTob(var pTob : Tob);
begin

  // création des champs de la ligne de tâche
  pTob.AddChampSup('PC_TACHE',true);
  pTob.AddChampSup('PC_PREVU',true);
  pTob.AddChampSup('PC_AFFECTE',true);
  pTob.AddChampSup('PC_REALISE',true);
  pTob.AddChampSup('PC_MOIS',true);
  pTob.AddChampSup('PC_MOIS1',true);
  pTob.AddChampSup('PC_MOIS2',true);
  pTob.AddChampSup('PC_MOIS3',true);
  pTob.AddChampSup('PC_MOIS4' ,true);
  pTob.AddChampSup('PC_MOIS5',true);
  pTob.AddChampSup('PC_MOIS6',true);
  pTob.AddChampSup('PC_MOIS7',true);
  pTob.AddChampSup('PC_MOIS8',true);
  pTob.AddChampSup('PC_MOIS9',true);
  pTob.AddChampSup('PC_MOIS10',true);
  pTob.AddChampSup('PC_MOIS11',true);
  pTob.AddChampSup('PC_RESTEAPASSERTHEORIQUE',true);
  pTob.AddChampSup('PC_RESTEAPASSER',true);
  pTob.AddChampSup('PC_TOTAL',true);
  pTob.AddChampSup('PC_ECART',true);

  // champ de gestion
  pTob.AddChampSup('PC_NUMEROTACHE',true);
  pTob.AddChampSup('PC_LIGNEFONCTION',true);
  pTob.AddChampSup('PC_RESSOURCE',true);
  pTob.AddChampSup('PC_FONCTION',true);
  pTob.AddChampSup('PC_AFFAIRE',true);
  pTob.AddChampSup('PC_ARTICLE',true);
  pTob.AddChampSup('PC_TIERS',true);

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : sauvegarde des qtes saisies dans le plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.SauvePlanDeCharge;
var
  i             : Integer;
  vBoResult     : Boolean;
  vTobLignesPla : Tob;
  vTobTaches    : Tob;
  vInNumligne   : Integer;
  vSt           : String;
  vQr           : TQuery;
begin

  InitMoveProgressForm (nil,'Traitement en cours...','',
                            fTobPlanCharge.detail.Count * 13, false, true) ;

  vTobLignesPla := Tob.create('Sauve PLanDeCharge', nil, -1);
  vTobTaches := Tob.create('TACHES', nil, -1);

  try
    Try
      // On raffraichit les données si elles n'ont pas ete raffraichies manuellement
      BRefreshClick(self);
      fTobPlanCharge.GetGridDetail(fGD_PlanCharge, fGD_PlanCharge.RowCount - 1, 'LePLanDeCharge', '');

      // suppression puis insert pour optimiser
      // on recherche / tache pour ne pas supprimer les taches de l'autre type (qte ou prix)
      for i := 0 to fTobPlanCharge.detail.count - 1 do
      begin
        if (i = 0) or (fTobPlanCharge.detail[i-1].getvalue('PC_NUMEROTACHE') <> fTobPlanCharge.detail[i].getvalue('PC_NUMEROTACHE')) then
          begin
            vSt := 'DELETE FROM AFPLANNING WHERE APL_AFFAIRE = "' + fStAffaire + '"';
            vSt := vSt + ' AND APL_NUMEROTACHE = "' + IntToStr(fTobPlanCharge.detail[i].getvalue('PC_NUMEROTACHE')) + '"';
            ExecuteSql(vSt);
          end;
      end;

      //comme on supprime, on redemarre au dernier de l'autre type
      vSt := 'select max(APL_NUMEROLIGNE) as nb from afplanning where apl_affaire = "' + fStAffaire + '"';
      vQr := openSql(vSt, True);
      try
        if Not vQR.Eof then vInNumLigne := vQR.findField('NB').AsInteger + 1 else vInNumLigne := 1;
      finally
        if vQR <> nil then ferme(vQR);
      end;

      // chargement des taches de l'affaire pour avoir l'ensemble des données des taches
      vBoResult := ChargeLesTaches (vTobTaches, fStAffaire, '');

      // on parcours a l'envers pour recuperer les qte des ressources
      // avant de valoriser la tache
      for i := fTobPlanCharge.detail.Count - 1 downto 0 do
        if vBoResult then
          vBoResult := PrepareSauvTob(vTobTaches, vTobLignesPla, fTobPlanCharge.Detail[i], vInNumLigne, nil)
        else
          break;

      if vBoResult then
        begin
          vTobTaches.UpdateDB(false);
          vTobLignesPla.insertDB(nil, false);
        end;

    except
      on E : Exception do
        begin
          MessageAlerte(TexteMsgTache[4] + #10#13 + E.Message);
          V_PGI.IOError := oeSaisie;
        end;
    end;

  Finally
    vTobLignesPla.Free;
    vTobTaches.Free;
    FiniMoveProgressForm;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Gestion des saisies float et recalcul des champs
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.fGD_PlanChargeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  fGD_PlanCharge.SynEnabled := false;

  // si la zone est vide (et que ce n'est pas une zone qui vient d'etre vidée)
  //ou est une colonne cachée
  if ((fGD_PlanCharge.Cells[ACol,ARow] = '') and
      (TobStrToFloat(fTobPlanCharge.Detail[ARow-1].GetValeur(1000 + aCol)) = 0)) or
     (fGD_PlanCharge.Cells[ACol,ARow] = #0) or
     (ACol = cInColTache) or (ACol = 21) or
     (ACol = 22) or (ACol = 23) or
     (ACol = 24) then

    cancel := false

  else
    begin

      // mettre 0 si null
      if (fGD_PlanCharge.Cells[ACol,ARow] = '') then
        fGD_PlanCharge.Cells[ACol,ARow] := '0';

      // si la colonne est le prévu
      if (ACol = cInColPrevu) then
        begin
          // Raf = Ancien Raf + nouveau prevu - ancien prevu
          fGD_PlanCharge.Cells[cInColRaf,ARow] := FloatToStr(valeur(fGD_PlanCharge.Cells[cInColRaf,ARow]) +
                                                  valeur(fGD_PlanCharge.Cells[cInColPrevu,ARow]) -
                                                  TobStrToFloat(fTobPlanCharge.Detail[ARow-1].GetValeur(1000 + cInColPrevu)));
                                                             
          if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) and (valeur(fGD_PlanCharge.Cells[cInColRaf,ARow]) < 0) then
            fGD_PlanCharge.Cells[cInColRaf,ARow] := '0';

          fGD_PlanCharge.Cells[cInColRafTh,ARow] := FloatToStr(valeur(fGD_PlanCharge.Cells[cInColRafTh,ARow]) +
                                                    valeur(fGD_PlanCharge.Cells[cInColPrevu,ARow]) -
                                                    TobStrToFloat(fTobPlanCharge.Detail[ARow-1].GetValeur(1000 + cInColPrevu)));

          if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) and (valeur(fGD_PlanCharge.Cells[cInColRafTh,ARow]) < 0) then
            fGD_PlanCharge.Cells[cInColRafTh,ARow] := '0';

          CalculTotal(aRow);
          CalculEcart(aRow);

        end

      // si la colonne est un mois
      else if (ACol > 3) and (ACol < 16) then
        begin
          // recalcul du reste a passer et de l'affecte

          // nouveau affecte = nouveau affecte mois - ancien affecte mois
          fGD_PlanCharge.Cells[cInColAffecte,ARow] := FloatToStr(valeur(fGD_PlanCharge.Cells[cInColAffecte,ARow]) +
                                                      valeur(fGD_PlanCharge.Cells[ACol,ARow]) -
                                                      TobStrToFloat(fTobPlanCharge.Detail[ARow-1].GetValeur(1000 + aCol)));

          // Raf = ancien Raf + ancien affecte mois - nouveau affecté mois
          fGD_PlanCharge.Cells[cInColRaf,ARow] :=
            FloatToStr(valeur(fGD_PlanCharge.Cells[cInColRaf,ARow]) -
                       valeur(fGD_PlanCharge.Cells[ACol,ARow]) +
                       TobStrToFloat(fTobPlanCharge.Detail[ARow-1].GetValeur(1000 + aCol)));

          if valeur(fGD_PlanCharge.Cells[cInColRaf,ARow]) < 0 then
            begin
              fGD_PlanCharge.Cells[cInColRaf,ARow] := '0';
              CalculTotal(aRow);
              CalculEcart(aRow);
            end;

          fGD_PlanCharge.Cells[cInColRafTh,ARow] :=
            FloatToStr(valeur(fGD_PlanCharge.Cells[cInColRafth,ARow]) -
                       valeur(fGD_PlanCharge.Cells[ACol,ARow]) +
                       TobStrToFloat(fTobPlanCharge.Detail[ARow-1].GetValeur(1000 + aCol)));
 
          if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) and (valeur(fGD_PlanCharge.Cells[cInColRafTh,ARow]) < 0) then
            fGD_PlanCharge.Cells[cInColRafTh,ARow] := '0';
        end

      // si la colonne est Raf
      else if aCol = cInColRaf then
        begin
          CalculTotal(aRow);
          CalculEcart(aRow);
        end;

      // controle de la valeur saisie
      if IsNumeric(fGD_PlanCharge.Cells[ACol,ARow]) then
        begin
          {if fBlGestionDec then  vInNbDec := 2 else vInNbDec := 0;
          vRdMontant := valeur(fGD_PlanCharge.Cells[ACol,ARow]);
          fGD_PlanCharge.Cells[ACol,ARow] := STRFMONTANT ( vRdMontant , 15 , vInNbDec, '' , true);
          }
          fGD_PlanCharge.Cells[ACol,ARow] := FormatJour(strToFloat(fGD_PlanCharge.Cells[ACol,ARow]));
          Cancel := False;
        end
      else
        Cancel := true;
    end;

  // transfert des données dans la tob
  fTobPlanCharge.GetGridDetail(fGD_PlanCharge, fGD_PlanCharge.RowCount - 1 , 'LePLanDeCharge', '');
  fGD_PlanCharge.SynEnabled := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : La ligne fonction en gras
Mots clefs ... :
*****************************************************************}
Procedure TOF_AFPLANCHARGE_ANCETRE.GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
begin
  // si ligne fonction
  if (fGD_PlanCharge.Cells[cInColLigFct,ARow] = 'X') then
    canvas.Font.Style := Canvas.Font.Style + [fsBold];
end;

Procedure TOF_AFPLANCHARGE_ANCETRE.GetCellCanvasTotaux ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
begin
  canvas.Font.Style := Canvas.Font.Style + [fsBold];
end;
   
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : Gestion de l'interdiction de saisir la Quantité initiale
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.fGD_PlanChargeSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  // test si on essaie d'aller sur une ligne fonction et sur la colonne qte initiale
  if (ACol = cInColPrevu) and (fGD_PlanCharge.Cells[cInColLigFct,ARow] = 'X') then
    CanSelect := False
  else
    CanSelect := True;

end;

procedure TOF_AFPLANCHARGE_ANCETRE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
    VK_ESCAPE : if Screen.ActiveControl = fGD_PlanCharge then
                  begin
                    key := 0;
                    fPEntete.SetFocus;
                  end
                else
                  Close;
    VK_RETURN : key := VK_TAB;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/05/2002
Modifié le ... :   /  /
Description .. : vérification de la cohérence de la saisie des
               : quantités prévues des taches
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.BVerifClick(Sender: TObject);
begin
  Verif(False);
end;

function TOF_AFPLANCHARGE_ANCETRE.Verif(pBoAsk : Boolean) : Boolean;
var
  i         : Integer;
  vRdTotal  : Double;
  vRdQte    : Double;
  vInTache  : Integer;
  vStTache  : String;
  vBoStop   : Boolean;

begin

  result := true;
  vBoStop := false;

  // pour la premiere tache
  vInTache := StrToInt(fGD_PlanCharge.Cells[cInColTache,1]);
  vStTache := fGD_PlanCharge.Cells[cInColNomTache,1];
  vRdTotal := valeur(fGD_PlanCharge.Cells[cInColPrevu,1]);
  vRdQte := 0;

  // pour chaque tache
  for i := 2 to fGD_PlanCharge.rowcount - 1 do
    begin
      if (vInTache = StrToInt(fGD_PlanCharge.Cells[cInColTache,i])) then
        begin
          vRdQte := vRdQte + valeur(fGD_PlanCharge.Cells[cInColPrevu,i]);
        end
      else
        // controle de l'egalite et changement de tache
        begin
          if (vRdQte <> vRdTotal) and (vRdQte <> 0) then
            begin
              //erreur uses DicoAF
              if pBoAsk then
                begin
                  if (PGIAskAF(format(TexteMsgTache[5],[vStTache]),'')= mrYes) then
                    begin
                      result := true;
                      vBoStop := true;
                    end
                  else
                    result := false;
                  Break;
                end
              else
                begin
                  PGIBoxAF (format(TexteMsgTache[1],[vStTache]),'');
                  result := false;
                  break;
                end;
            end
          else
            begin
              vRdQte := 0;
              vRdTotal := valeur(fGD_PlanCharge.Cells[cInColPrevu,i]);
              vInTache := StrToInt(fGD_PlanCharge.Cells[cInColTache,i]);
              vStTache := fGD_PlanCharge.Cells[cInColNomTache,i];
            end;
        end;
    end;

  // controle de l'egalite pour la derniere tache
  if (vRdQte <> vRdTotal) and (vRdQte <> 0) and ((result) and (not vBoStop)) then
    begin

      if pBoAsk then
        begin
          if (PGIAskAF(format(TexteMsgTache[5],[vStTache]),'') = mrYes) then
            result := true
          else
             result := false;
        end
      else
        begin
          PGIBoxAF (format(TexteMsgTache[1],[vStTache]),'');
          result := false;
        end;
    end;

  if (result) and (not pBoAsk) then PGIInfoAF (TexteMsgTache[2],'');

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/05/2002
Modifié le ... :   /  /
Description .. : Recalcul des quantités des lignes fonctions
               :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.BRefreshClick(Sender: TObject);
var
  i,j           : Integer;
  vInTache      : Integer;
  vInIndiceFct  : Integer;
  vInLastIndice : Integer;
  vRdMois       : array [0..11] of Double;
  vRdQteAffecte : Double;
  vRdQteRaf     : Double;
  vRdQteRafTh   : Double;

  procedure CalculLigne(vInIndice : Integer; vRdQteRaf, vRdQteRafTh, vRdQteAffecte : Double; vRdMois : Array of double);
  var
    j : integer;
  begin

    // Raf = Somme des Raf des ressources - le réalisé de la fonction
    fGD_PlanCharge.Cells[cInColRaf, vInIndice] := FormatJour(vRdQteRaf - valeur(fGD_PlanCharge.Cells[cInColRealise, vInIndice]));
    fGD_PlanCharge.Cells[cInColRafTh, vInIndice] := FormatJour(vRdQteRafTh - valeur(fGD_PlanCharge.Cells[cInColRealise, vInIndice]));

    if valeur(fGD_PlanCharge.Cells[cInColRaf, vInIndice]) < 0 then fGD_PlanCharge.Cells[cInColRaf, vInIndice] := '0';
    if valeur(fGD_PlanCharge.Cells[cInColRafTh, vInIndice]) < 0 then fGD_PlanCharge.Cells[cInColRafTh, vInIndice] := '0';

    // colonne affectée                                     
    fGD_PlanCharge.Cells[cInColAffecte, vInIndice] := FormatJour(vRdQteAffecte);

    CalculTotal(vInIndice);
    CalculEcart(vInIndice);

    // mois
    for j := 0 to 11 do
      if vRdMois[j] = 0 then
        fGD_PlanCharge.Cells[cInColDebMois + j , vInIndice] := ''
      else
        fGD_PlanCharge.Cells[cInColDebMois + j , vInIndice] := FormatJour(vRdMois[j]);

  end;

begin

  // changement de cellule pour executer le cellExit
  fGD_PlanCharge.Col := cInColAffecte;

  // pour la premiere tache
  vInTache      := StrToInt(fGD_PlanCharge.Cells[cInColTache,1]);
  vInIndiceFct  := 1;
  vInLastIndice := 2;
  vRdQteAffecte := 0;
  vRdQteRafTh   := 0;
  vRdQteRaf     := 0;

  for j := 0 to 11 do vRdMois[j] := 0;

  // pour chaque tache
  for i := 2 to fGD_PlanCharge.rowcount - 1 do
    begin
      // si on est toujours sur la meme tache
      if (vInTache = StrToInt(fGD_PlanCharge.Cells[cInColTache,i])) then
        begin
          vRdQteAffecte := vRdQteAffecte + TobStrToFloat(fGD_PlanCharge.Cells[cInColAffecte,i]);
          vRdQteRaf     := vRdQteRaf     + TobStrToFloat(fGD_PlanCharge.Cells[cInColRaf,i]) + TobStrToFloat(fGD_PlanCharge.Cells[cInColRealise,i]);
          vRdQteRafTh   := vRdQteRafTh   + TobStrToFloat(fGD_PlanCharge.Cells[cInColRafTh,i]) + TobStrToFloat(fGD_PlanCharge.Cells[cInColRealise,i]);

          if vRdQteRaf < 0 then vRdQteRaf := 0;
          if vRdQteRafTh < 0 then vRdQteRafTh := 0;

          for j := 0 to 11 do
            vRdMois[j] := vRdMois[j] + TobStrToFloat(fGD_PlanCharge.Cells[cInColDebMois + j ,i]);
        end
      else
        // affectation des valeurs et changement de tache
        begin
          // pas de calcul de l'affecte sur la ligne maitre si pas de ligne detail
          if i <> (vInIndiceFct + 1) then
            CalculLigne(vInIndiceFct, vRdQteRaf, vRdQteRafTh, vRdQteAffecte, vRdMois);

          vInTache      := StrToInt(fGD_PlanCharge.Cells[cInColTache,i]);
          vRdQteAffecte := 0;
          vRdQteRafTh   := 0;
          vRdQteRaf     := 0;
          vInIndiceFct  := i;
          for j := 0 to 11 do vRdMois[j] := 0;
        end;
      vInLastIndice := i;
    end;

  // on ne fait rien si une seule ligne fonction
  // pour la derniere fonction
  // pas de calcul de l'affecte sur la ligne maitre si pas de ligne detail
  If (fGD_PlanCharge.rowcount > 2) and
     (vInLastIndice <> vInIndiceFct) and
     (vInLastIndice >= vInIndiceFct) then
    CalculLigne(vInIndiceFct, vRdQteRaf, vRdQteRafTh, vRdQteAffecte, vRdMois);

  LoadTotal;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : Enregistrement en base
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.BValiderClick(Sender: TObject);
begin
  if Verif(True) then
    begin
      transactions(SauvePlanDeCharge, 1);
      ecran.ModalResult := MrYes;
    end
  else
    ecran.ModalResult := MrNone;
end;                                          

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : impression de la grid
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.BImprimerClick(Sender: TObject);
//var
//  vSt : String;
begin  
lanceEtattob ('E','APL','APC',fTobPlanCharge,true,false,false,nil,'','Edition Plan de charge',false);
//  vSt := CodeAffaireAffiche(fStAffaire,' ');
//  PrintDBGrid(fGD_PlanCharge, nil, 'Plan de Charge / Affaire : ' + vSt + '   ' + fStLibAffaire, '') ;
end;
  
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : formatage de la grille du plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_ANCETRE.FormatPCGrid;
var
  i : Integer;
  j : Integer;
  
begin

  fGD_PlanCharge.rowcount := fTobPlanCharge.Detail.count + 1;
  fGD_PlanCharge.colcount := 22;
  fGD_PlanCharge.CellValues[0,0] := '';

  fGD_PlanCharge.CellValues[4,0]  := MoisCourant(0);
  fGD_PlanCharge.CellValues[5,0]  := MoisCourant(1);
  fGD_PlanCharge.CellValues[6,0]  := MoisCourant(2);
  fGD_PlanCharge.CellValues[7,0]  := MoisCourant(3);
  fGD_PlanCharge.CellValues[8,0]  := MoisCourant(4);
  fGD_PlanCharge.CellValues[9,0]  := MoisCourant(5);
  fGD_PlanCharge.CellValues[10,0] := MoisCourant(6);
  fGD_PlanCharge.CellValues[11,0] := MoisCourant(7);
  fGD_PlanCharge.CellValues[12,0] := MoisCourant(8);
  fGD_PlanCharge.CellValues[13,0] := MoisCourant(9);
  fGD_PlanCharge.CellValues[14,0] := MoisCourant(10);
  fGD_PlanCharge.CellValues[15,0] := MoisCourant(11);

  // colonnes cachées
  fGD_PlanCharge.ColWidths[cInColTache] := -1; fGD_PlanCharge.ColWidths[cInColLigFct] := -1;
  fGD_PlanCharge.ColLengths[cInColTache] := -1; fGD_PlanCharge.ColLengths[cInColLigFct] := -1;

  // lignes
  for i := 0 to fGD_PlanCharge.rowCount - 1 do
    fGD_PlanCharge.RowHeights[i] := 17;

  fGD_PlanCharge.RowHeights[0] := 22;

  // alignement
  for i := 1 to fGD_PlanCharge.ColCount -1 do
    begin
      fGD_PlanCharge.ColAligns[i] := TaRightJustify;
      fGD_PlanCharge.ColTypes[i] := 'R';
    end;

  // arrondi
  // C.B 11/08/03
  for i := 1 to 19 do
    for j := 1 to fGD_PlanCharge.RowCount -1 do
      if fGD_PlanCharge.Cells[i,j] <> #0 then
        fGD_PlanCharge.Cells[i,j] := FormatJour(strToFloat(fGD_PlanCharge.Cells[i,j]));
 
  // colonnes en lecture seule
  fGD_PlanCharge.ColEditables[cInColAffecte]  := False;
  fGD_PlanCharge.ColEditables[cInColRealise]  := False;
  fGD_PlanCharge.ColEditables[cInColRafTh]    := False;
  fGD_PlanCharge.ColEditables[cInColTotal]    := False;
  fGD_PlanCharge.ColEditables[cInColEcart]    := False;

  fGD_PlanCharge.ColEditables[cInColTache]    := False;
  fGD_PlanCharge.ColEditables[cInColLigFct]   := False;
   
  // format grille totaux
  fGD_TOTAUX.GetCellCanvas := GetCellCanvasTotaux;
  fGD_TOTAUX.rowcount := 1;
  fGD_TOTAUX.RowHeights[0] := 22;
  fGD_TOTAUX.colcount := 22;

end;

procedure TOF_AFPLANCHARGE_ANCETRE.CalculTotal(aRow : Integer);
Begin
  // total = affecte + realise + Raf
  fGD_PlanCharge.Cells[cInColTotal,ARow] := FormatJour(valeur(fGD_PlanCharge.Cells[cInColRealise,ARow]) +
                                            valeur(fGD_PlanCharge.Cells[cInColAffecte,ARow]) +
                                            valeur(fGD_PlanCharge.Cells[cInColRaf,ARow]));
End;

procedure TOF_AFPLANCHARGE_ANCETRE.CalculEcart(aRow : Integer);
Begin
  // ecart = Total - Prevu
  fGD_PlanCharge.Cells[cInColEcart,ARow] := FormatJour(valeur(fGD_PlanCharge.Cells[cInColTotal,ARow]) -
                                            valeur(fGD_PlanCharge.Cells[cInColPrevu,ARow]));
end;

{******************************************************************************}
{* TOF_AFPLANCHARGE_QTE                                                       *}
{******************************************************************************}
procedure TOF_AFPLANCHARGE_QTE.LoadTaches;
var
  vQr               : TQuery;
  S                 : String;
  i                 : integer;
  vTobGlobale       : Tob;

begin

  // creation de la tob de travail
  fTobPlanCharge := TOB.create('Le PLanDeCharge', nil, -1);
  FormatTob(fTobPlanCharge);

  vTobGlobale := TOB.create('Les Taches', nil, -1);

  // Chargement des données globales pour chaques taches
  S := 'select ATA_LIBELLETACHE1, ATA_QTEINITIALE, ATA_FONCTION, ATA_FONCTION AS ATR_FONCTION,';
  S := S + ' ATA_QTEAPLANIFIER, ATA_NUMEROTACHE, ATA_ARTICLE, ATA_ARTICLE AS ATR_ARTICLE, ATA_TIERS ';
  S := S + ' FROM AFFAIRE, TACHE ';
  S := S + ' where AFF_AFFAIRE = "' + fStAffaire + '"';
  S := S + ' and ATA_MODESAISIEPDC = "QUA"';
  S := S + ' and AFF_AFFAIRE = ATA_AFFAIRE ';
  S := S + ' ORDER BY ATA_NUMEROTACHE';


  vQr := nil;
  Try
    vQr := OpenSql(S, True);
    if Not vQr.Eof then
      begin
        vTobGlobale.LoadDetailDB('LesTaches','','',vQr,False,True);
        vTobGlobale.AddChampSup('ATA_QTEAPLANIFIERCALC', true);
        vTobGlobale.AddChampSup('ATA_RAPPTPRCAL', true);

        // pour chaque fonction de chaque tache
        for i := 0 to vTobGlobale.Detail.Count - 1  do
          begin

            // Ajout de la ligne globale (fonction)
            LoadFonction(vTobGlobale.detail[i]);
                           
            // recherche des ressources qui ont de l'activite mais qui
            // ne sont pas prévues
            if GetParamSoc('SO_AFCLIENT') = cInClientAlgoe then
              RechercheRes(vTobGlobale.detail[i]);
                                    
            // Ajout des lignes des ressources
            LoadRessources(vTobGlobale.detail[i]);

          end;

        // Chargement de la grille
        fTobPlanCharge.PutGridDetail(fGD_PlanCharge, True, False, '');
        FormatPCGrid;

        // positionnement du curseur
        fGD_PlanCharge.col := 4;
      end;

  Finally
    if vQr <> Nil then Ferme(vQr);
    vTobGlobale.Free;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Chargement des données de chaque tâches
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_QTE.LoadFonction(pTobTache : Tob);
var
  vTobLigne       : Tob;
  vRdQteAffecte   : Double;
  vRdQteRealise   : Double;
  vStQteInitiale  : String;
  vRdQteRAP       : Double;

begin

  vTobLigne := TOB.create('Une Ligne', fTobPlanCharge, -1);
  FormatTob(vTobLigne);
  vStQteInitiale := FormatJour(pTobTache.GetValue('ATA_QTEINITIALE'));

  // Chargement des quantites des mois de la ligne
  vRdQteAffecte := LoadMois(vTobLigne, pTobTache, True, fStAffaire, 'ATA', 'APL_QTEPLANIFIEE');

  // Calcul de la quantité réalisée pour la fonction
  vRdQteRealise := CalculRealise(fStAffaire, pTobTache.GetValue('ATA_FONCTION'), 'ACT_QTE','', pTobTache.GetValue('ATA_ARTICLE'));

  // remplir les quantites globales de la tache
  vTobLigne.putValue('PC_TACHE', pTobTache.GetValue('ATA_LIBELLETACHE1'));
  vTobLigne.putValue('PC_PREVU', vStQteInitiale);
  vTobLigne.putValue('PC_AFFECTE', FormatJour(vRdQteAffecte));
  vTobLigne.putValue('PC_REALISE', FormatJour(vRdQteRealise));

  vRdQteRAP := valeur(vStQteInitiale) - vRdQteAffecte - vRdQteRealise + valeur(pTobTache.GetValue('ATA_QTEAPLANIFIER'));
  vTobLigne.putValue('PC_RESTEAPASSERTHEORIQUE', FormatJour(vRdQteRAP));
  vTobLigne.putValue('PC_RESTEAPASSER', FormatJour(vRdQteRAP));

  vTobLigne.putValue('PC_TOTAL', FormatJour(vRdQteAffecte +
                      valeur(vTobLigne.GetValue('PC_RESTEAPASSER')) +
                      vRdQteRealise));

  vTobLigne.putValue('PC_ECART', FormatJour(valeur(vTobLigne.GetValue('PC_TOTAL')) -
                                 valeur(vTobLigne.GetValue('PC_PREVU'))));

  // Champs de gestion
  vTobLigne.putValue('PC_NUMEROTACHE', pTobTache.GetValue('ATA_NUMEROTACHE'));
  vTobLigne.putValue('PC_LIGNEFONCTION', 'X');
  vTobLigne.putValue('PC_RESSOURCE', '');
  vTobLigne.putValue('PC_FONCTION', pTobTache.GetValue('ATA_FONCTION'));
  vTobLigne.putValue('PC_AFFAIRE', fStAffaire);
  vTobLigne.putValue('PC_ARTICLE', pTobTache.GetValue('ATA_ARTICLE'));
  vTobLigne.putValue('PC_TIERS', pTobTache.GetValue('ATA_TIERS'));

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Chargement des lignes de ressources
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_QTE.LoadRessources(pTobTache : Tob);
var
  i               : Integer;
  S               : String;
  vQrRes          : TQuery;
  vTobRes         : Tob;
  vTobLigne       : Tob;
  vRdQteAffecte   : Double;
  vRdQteRealise   : Double;
  vStQteInitiale  : String;
  vRdQteRAP       : Double;

begin

  vTobRes := TOB.create('Les Ressources', nil, -1);

  // Selection des données globales pour chaque ressources
  S := 'select ATR_RESSOURCE, ATR_FONCTION, ARS_LIBELLE, ATR_QTEINITIALE, ATA_ARTICLE AS ATR_ARTICLE, ';
  S := S + ' ATR_QTEAPLANIFIER, ATR_NUMEROTACHE '; // ATR_QTEAFFECTE
  S := S + ' FROM TACHE, TACHERESSOURCE, RESSOURCE ';
  S := S + ' where ATA_AFFAIRE = "' + fStAffaire + '"';
  S := S + ' and ATA_NUMEROTACHE = ' + IntToStr(pTobTache.Getvalue('ATA_NUMEROTACHE'));
  S := S + ' and ATA_NUMEROTACHE = ATR_NUMEROTACHE ';
  S := S + ' and ATA_AFFAIRE = ATR_AFFAIRE ';
  S := S + ' and ARS_RESSOURCE = ATR_RESSOURCE ';


  vQrRes := nil;
  Try
    vQrRes := OpenSql(S,True);
    if Not vQrRes.Eof then
      begin
        vTobRes.LoadDetailDB('LesRessources','','',vQrRes,False,True);

        // pour chaque ressource
        for i := 0 to vTobRes.Detail.Count - 1 do
          begin

            if vTobRes.Detail[i].GetValue('ATR_QTEINITIALE') = NULL then vTobRes.Detail[i].putValue('ATR_QTEINITIALE','0');
            if vTobRes.Detail[i].GetValue('ATR_QTEAPLANIFIER') = NULL then vTobRes.Detail[i].putValue('ATR_QTEAPLANIFIER', '0');

            vTobLigne := TOB.create('Une Ligne', fTobPlanCharge, -1);
            FormatTob(vTobLigne);

            // Chargement des quantites des mois de la ligne
            vRdQteAffecte := LoadMois(vTobLigne, vTobRes.Detail[i], False, fStAffaire, 'ATR', 'APL_QTEPLANIFIEE');
            vStQteInitiale := FormatJour(vTobRes.Detail[i].GetValue('ATR_QTEINITIALE'));

            // Calcul de la quantité réalisée pour la ressource
            vRdQteRealise := CalculRealise(fStAffaire, pTobTache.GetValue('ATA_FONCTION'), 'ACT_QTE', vTobRes.Detail[i].GetValue('ATR_RESSOURCE'), pTobTache.GetValue('ATA_ARTICLE'));

            // remplir les quantites globales de la Ressource
            vTobLigne.putValue('PC_TACHE', vTobRes.Detail[i].GetValue('ARS_LIBELLE'));
            vTobLigne.putValue('PC_PREVU', vStQteInitiale);
            vTobLigne.putValue('PC_AFFECTE', FormatJour(vRdQteAffecte));
            vTobLigne.putValue('PC_REALISE', FormatJour(vRdQteRealise));

            vRdQteRAP := valeur(vStQteInitiale) - vRdQteRealise - vRdQteAffecte + valeur(vTobRes.Detail[i].GetValue('ATR_QTEAPLANIFIER'));
            vTobLigne.putValue('PC_RESTEAPASSERTHEORIQUE', FormatJour(vRdQteRAP));
            vTobLigne.putValue('PC_RESTEAPASSER', FormatJour(vRdQteRAP));

            vTobLigne.putValue('PC_TOTAL',FormatJour(vRdQteAffecte +
                                          valeur(vTobLigne.GetValue('PC_RESTEAPASSER')) +
                                          vRdQteRealise));

            vTobLigne.putValue('PC_ECART', FormatJour(valeur(vTobLigne.GetValue('PC_TOTAL')) -
                                           valeur(vTobLigne.GetValue('PC_PREVU'))));

            // Champs de gestion
            vTobLigne.putValue('PC_NUMEROTACHE', vTobRes.Detail[i].GetValue('ATR_NUMEROTACHE'));
            vTobLigne.putValue('PC_LIGNEFONCTION', '');
            vTobLigne.putValue('PC_RESSOURCE', vTobRes.Detail[i].GetValue('ATR_RESSOURCE'));
            vTobLigne.putValue('PC_FONCTION', vTobRes.Detail[i].GetValue('ATR_FONCTION'));
            vTobLigne.putValue('PC_AFFAIRE', fStAffaire);
            vTobLigne.putValue('PC_ARTICLE', vTobRes.Detail[i].GetValue('ATR_ARTICLE'));
            vTobLigne.putValue('PC_TIERS', pTobTache.GetValue('ATA_TIERS'));
          end;
      end;
  finally
    if vQrRes <> nil then Ferme(vQrRes);
    vTobRes.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : formatage de la grille du plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_QTE.FormatPCGrid;
begin
  inherited;
   fGD_PlanCharge.ColWidths[0] := 59;

  fGD_PlanCharge.CellValues[cInColPrevu,0] := TraduitGA('Prévu');
  fGD_PlanCharge.CellValues[cInColAffecte,0] := TraduitGA('Affect.');
  fGD_PlanCharge.CellValues[cInColRealise,0] := TraduitGA('Réal.');
  fGD_PlanCharge.CellValues[cInColRafTh,0] := TraduitGA('RAFTh');
  fGD_PlanCharge.CellValues[cInColRaf,0]   := TraduitGA('RAF');
  fGD_PlanCharge.CellValues[cInColTotal,0] := TraduitGA('Total');
  fGD_PlanCharge.CellValues[cInColEcart,0] := TraduitGA('Ecart');

  fGD_PlanCharge.ColWidths[cInColPrevu] := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColAffecte] := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColRealise] := cInTitre2;

  fGD_PlanCharge.ColWidths[4] := cInMois; fGD_PlanCharge.ColWidths[5] := cInMois;
  fGD_PlanCharge.ColWidths[6] := cInMois; fGD_PlanCharge.ColWidths[7] := cInMois;
  fGD_PlanCharge.ColWidths[8] := cInMois; fGD_PlanCharge.ColWidths[9] := cInMois;
  fGD_PlanCharge.ColWidths[10] := cInMois; fGD_PlanCharge.ColWidths[11] := cInMois;
  fGD_PlanCharge.ColWidths[12] := cInMois; fGD_PlanCharge.ColWidths[13] := cInMois;
  fGD_PlanCharge.ColWidths[14] := cInMois; fGD_PlanCharge.ColWidths[15] := cInMois;

  fGD_PlanCharge.ColWidths[cInColRafTh] := cInTitre2; fGD_PlanCharge.ColWidths[cInColRaf]   := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColTotal] := cInTitre2; fGD_PlanCharge.ColWidths[cInColEcart] := cInTitre3;

  TFVierge(Ecran).HMTrad.ResizeGridColumns(fGD_PlanCharge);
end;                                                          

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : sauvegarde d'une ligne de planning
Mots clefs ... :
*****************************************************************}
function TOF_AFPLANCHARGE_QTE.PrepareSauvTob(pTobTaches, pTobLignesPla, pTobLigne : Tob; var pInNumLigne : Integer; TQR : TQRProgressForm) : Boolean;
var
  vTob              : Tob;
  vRdQte            : Double;
  vDtLigne          : TDateTime;
  i                 : Integer;
  vStAff0, vStAff1  : String;
  vStAff2, vStAff3  : String;
  vStAvenant        : String;
  vStTiers          : String;
  vArticle          : RecordArticle;
  vTobTache         : Tob;
  vTobRes           : Tob;
  vRdQteAPlanif     : Double;
  vDtFin            : TDateTime;
  vTOBAffaires      : Tob;
  vTobArticles      : Tob;
  vAFOAssistants    : TAFO_Ressources;

begin

  Result := True;
  MoveCurProgressForm ('');

  vAFOAssistants  := TAFO_Ressources.Create;
  vTOBAffaires    := TOB.Create('Les Affaires',Nil,-1) ;
  vTOBArticles    := TOB.Create('Les Articles',Nil,-1) ;

  try
    if pTobLigne.GetValue('PC_LIGNEFONCTION') = 'X' then
      begin
        // on se positionne sur la tache
        vTobTache := pTobTaches.FindFirst(['ATA_AFFAIRE', 'ATA_NUMEROTACHE'],[fStAffaire, pTobLigne.GetValue('PC_NUMEROTACHE')], False);

        // on ne sauvegarde plus que le delta
        //vTobTache.PutValue('ATA_QTEAPLANIFIER',pTobLigne.GetValue('PC_RESTEAPASSER'));
        vRdQteAPlanif := valeur(pTobLigne.GetValue('PC_RESTEAPASSER')) -
                         valeur(pTobLigne.GetValue('PC_PREVU')) +
                         valeur(pTobLigne.GetValue('PC_AFFECTE')) +
                         valeur(pTobLigne.GetValue('PC_REALISE'));

        vTobTache.PutValue('ATA_QTEAPLANIFIER', floatToStr(vRdQteAPlanif));
        vTobTache.PutValue('ATA_QTEAPLANIFUREF', floatToStr(conversionUnite(vTobTache.GetValue('ATA_UNITETEMPS'),
                                                                            getparamsoc('SO_AFMESUREACTIVITE'),
                                                                            vRdQteAPlanif)));

        // valorisation des taches et des ressources
        Valorisation(vTobTache, 'ATA', vAFOAssistants, vTOBAffaires, vTobArticles);
      end
    else
      begin
        // on se positionne sur la ressource
        vTobRes := pTobTaches.FindFirst(['ATR_AFFAIRE', 'ATR_NUMEROTACHE', 'ATR_RESSOURCE'],[fStAffaire, pTobLigne.GetValue('PC_NUMEROTACHE'), pTobLigne.GetValue('PC_RESSOURCE')], True);
        // on est en unité jour dans la tache pour le plan de charge
        vTobRes.PutValue('ATR_QTEINITIALE', pTobLigne.GetValue('PC_PREVU'));
        vTobRes.PutValue('ATR_QTEINITUREF', floatToStr(conversionUnite('J',
                                                                       getparamsoc('SO_AFMESUREACTIVITE'),
                                                                       valeur(pTobLigne.GetValue('PC_PREVU')))));

        //vTobRes.PutValue('ATR_QTEAPLANIFIER', pTobLigne.GetValue('PC_RESTEAPASSER'));
        vRdQteAPlanif := valeur(pTobLigne.GetValue('PC_RESTEAPASSER')) -
                         valeur(pTobLigne.GetValue('PC_PREVU')) +
                         valeur(pTobLigne.GetValue('PC_AFFECTE')) +
                         valeur(pTobLigne.GetValue('PC_REALISE'));

        vTobRes.PutValue('ATR_QTEAPLANIFIER', floatToStr(vRdQteAPlanif));
        vTobRes.PutValue('ATR_QTEAPLANIFUREF', floatToStr(conversionUnite('J',
                                                                       getparamsoc('SO_AFMESUREACTIVITE'),
                                                                       vRdQteAPlanif)));

        // la valorisation des ressources est faites en meme temps que la tache
        // Valorisation(vTobRes, 'ATR', True);
      end;

    CodeAffaireDecoupe(fStAffaire, vStAff0, vStAff1, vStAff2, vStAff3,vStAvenant, taModif, false);

    // C.B prise en compte des affaires modeles
    if TeststCleAffaire(fStAffaire,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant,vStTiers, False, True, False, False) <> 1 then
      vStTiers := '';
    GetArticle(pTobLigne.GetValue('PC_ARTICLE'), vArticle);

    // parcours de la ligne du plan de charge
    for i := cInColDebMois to cInColFinMois do
      begin
        MoveCurProgressForm ('');
        vDtLigne := CalculDate(i-cInColDebMois);
        vRdQte := TobStrToFloat(pTobLigne.GetValeur(1000 + i));

        // on insere un enregistrement que si qte est non nulle
        if vRdQte <> 0 then
          Begin
            vTob := Tob.create('AFPLANNING', pTobLignesPla, -1);
            Try
              vTob.PutValue('APL_AFFAIRE', fStAffaire);
              vTob.PutValue('APL_NUMEROLIGNE', pInNumLigne);
              vTob.PutValue('APL_TYPELIGNEPLA', 'TAC');
              vTob.PutValue('APL_NUMEROTACHE', pTobLigne.GetValue('PC_NUMEROTACHE'));
              vTob.PutValue('APL_RESSOURCE', pTobLigne.GetValue('PC_RESSOURCE'));
              vTob.PutValue('APL_FONCTION', pTobLigne.GetValue('PC_FONCTION'));
              vTob.PutValue('APL_ARTICLE', pTobLigne.GetValue('PC_ARTICLE'));
              vTob.PutValue('APL_TYPEARTICLE', vArticle.StTypeArticle);
              vTob.PutValue('APL_CODEARTICLE', vArticle.StCodeArticle);
              vTob.PutValue('APL_TIERS', vStTiers);
              vTob.PutValue('APL_ETATLIGNE', getparamsoc('SO_AFPLANNINGETAT'));
              vTob.PutValue('APL_TERMINE', '-');
              vTob.PutValue('APL_LIGNEGENEREE', '-');
              vTob.PutValue('APL_ACTIVITEGENERE', '-');
              vTob.PutValue('APL_AFFAIRE0', vStAff0);
              vTob.PutValue('APL_AFFAIRE1', vStAff1);
              vTob.PutValue('APL_AFFAIRE2', vStAff2);
              vTob.PutValue('APL_AFFAIRE3', vStAff3);
              vTob.PutValue('APL_AVENANT', vStAvenant);
              vTob.PutValue('APL_ACTIVITEEFFECT','-');

              // pour le plan de charge, on est en unité jour
              vTob.PutValue('APL_UNITETEMPS', 'J');
              vTob.PutValue('APL_QTEPLANIFIEE', FloatToStr(vRdQte));
              vTob.PutValue('APL_QTEPLANIFUREF', floatToStr(conversionUnite('J', GetParamSoc('SO_AFMESUREACTIVITE'), vRdQte)));
              vTob.PutValue('APL_QTEREALISE', FloatToStr(0));
              vTob.PutValue('APL_QTEREALUREF', FloatToStr(0));
              vTob.PutValue('APL_DATEDEBPLA', vDtLigne);

              // date fin = date deb + qte
              vDtFin := plusDate(vDtLigne, round(vRdQte - 1), 'J');
              vTob.PutValue('APL_DATEFINPLA', vDtFin);
              //vTob.PutValue('APL_DATEFINPLA', iDate2099);
              vTob.PutValue('APL_LIBELLEPLA', pTobLigne.GetValue('PC_TACHE'));

              // valorisation des lignes de planning
              Valorisation(vTob, 'APL', vAFOAssistants, vTOBAffaires, vTobArticles);
              pInNumLigne := pInNumLigne + 1;
            Except
              PGIBoxAF (TexteMsgTache[3],'');
              result := false;
            End;
          end;
      End;
  finally
    vTOBAffaires.Free;
    vTobArticles.Free;
    vAFOAssistants.Free;
  end;
end;

{******************************************************************************}
{* TOF_AFPLANCHARGE_MNT                                                       *}
{******************************************************************************}
procedure TOF_AFPLANCHARGE_MNT.LoadTaches;
var
  vQr               : TQuery;
  S                 : String;
  i                 : integer;
  vTobGlobale       : Tob;

begin

  // creation de la tob de travail
  fTobPlanCharge := TOB.create('Le PLanDeCharge', nil, -1);
  FormatTob(fTobPlanCharge);

  vTobGlobale := TOB.create('Les Taches', nil, -1);

  // Chargement des données globales pour chaques taches
  S := 'select ATA_LIBELLETACHE1, ATA_INITPTPR, ATA_FONCTION, ATA_FONCTION AS ATR_FONCTION,';
  S := S + ' ATA_RAPPTPR, ATA_NUMEROTACHE, ATA_ARTICLE, ATA_ARTICLE AS ATR_ARTICLE, ATA_TIERS ';
  S := S + ' FROM AFFAIRE, TACHE ';
  S := S + ' where AFF_AFFAIRE = "' + fStAffaire + '"';
  S := S + ' and ATA_MODESAISIEPDC = "MPR"';
  S := S + ' and AFF_AFFAIRE = ATA_AFFAIRE ';
  S := S + ' ORDER BY ATA_NUMEROTACHE';

  vQr := nil;
  Try
    vQr := OpenSql(S, True);
    if Not vQr.Eof then
      begin

        vTobGlobale.LoadDetailDB('LesTaches','','',vQr,False,True);
        vTobGlobale.AddChampSup('ATA_RAPPTPRCALC', true);
        vTobGlobale.AddChampSup('ATA_RAPPTPRCAL', true);

        // pour chaque fonction de chaque tache
        for i := 0 to vTobGlobale.Detail.Count - 1  do

          // Ajout de la ligne globale (fonction)
          LoadFonction(vTobGlobale.detail[i]);

        // Chargement de la grille
        fTobPlanCharge.PutGridDetail(fGD_PlanCharge, True, False, '');
        FormatPCGrid;                                                  

        // positionnement du curseur
        fGD_PlanCharge.col := 4;
      end;

  Finally
    if vQr <> Nil then Ferme(vQr);
    vTobGlobale.Free;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Chargement des données de chaque tâches
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_MNT.LoadFonction(pTobTache : Tob);
var
  vTobLigne       : Tob;
  vRdQteAffecte   : Double;
  vRdQteRealise   : Double;
  vStQteInitiale  : String;
  vRdQteRAP       : Double;

begin

  vTobLigne := TOB.create('Une Ligne', fTobPlanCharge, -1);
  FormatTob(vTobLigne);
  vStQteInitiale := FormatJour(pTobTache.GetValue('ATA_INITPTPR'));

  // Chargement des quantites des mois de la ligne
  vRdQteAffecte := LoadMois(vTobLigne, pTobTache, True, fStAffaire, 'ATA', 'APL_INITPTPR');

  // Calcul de la quantité réalisée pour la fonction
  vRdQteRealise := CalculRealise(fStAffaire, '', 'ACT_TOTPR', '', pTobTache.GetValue('ATA_ARTICLE'));
                                              
  // remplir les quantites globales de la tache
  vTobLigne.putValue('PC_TACHE', pTobTache.GetValue('ATA_LIBELLETACHE1'));
  vTobLigne.putValue('PC_PREVU', vStQteInitiale);
  vTobLigne.putValue('PC_AFFECTE', FormatJour(vRdQteAffecte));
  vTobLigne.putValue('PC_REALISE', FormatJour(vRdQteRealise));

  vRdQteRAP := valeur(vStQteInitiale) - vRdQteAffecte - vRdQteRealise + valeur(pTobTache.GetValue('ATA_RAPPTPR'));
  vTobLigne.putValue('PC_RESTEAPASSERTHEORIQUE', FormatJour(vRdQteRAP));
  vTobLigne.putValue('PC_RESTEAPASSER', FormatJour(vRdQteRAP));

  vTobLigne.putValue('PC_TOTAL', FormatJour(vRdQteAffecte +
                      valeur(vTobLigne.GetValue('PC_RESTEAPASSER')) +
                      vRdQteRealise));

  vTobLigne.putValue('PC_ECART', FormatJour(valeur(vTobLigne.GetValue('PC_TOTAL')) -
                                 valeur(vTobLigne.GetValue('PC_PREVU'))));

  // Champs de gestion
  vTobLigne.putValue('PC_NUMEROTACHE', pTobTache.GetValue('ATA_NUMEROTACHE'));
  vTobLigne.putValue('PC_LIGNEFONCTION', 'X');
  vTobLigne.putValue('PC_RESSOURCE', '');
  vTobLigne.putValue('PC_FONCTION', pTobTache.GetValue('ATA_FONCTION'));
  vTobLigne.putValue('PC_AFFAIRE', fStAffaire);
  vTobLigne.putValue('PC_ARTICLE', pTobTache.GetValue('ATA_ARTICLE'));
  vTobLigne.putValue('PC_TIERS', pTobTache.GetValue('ATA_TIERS'));  

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : formatage de la grille du plan de charge en montant
                 pour les colonnes, on augmente la taille prévue pour
                 les quantités, car elle n'est pas suffisante pour
                 des montants
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANCHARGE_MNT.FormatPCGrid;
begin
  inherited;
  fGD_PlanCharge.ColWidths[0] := 59;
  fGD_PlanCharge.CellValues[cInColPrevu,0] := TraduitGA('Prévu');
  fGD_PlanCharge.CellValues[cInColAffecte,0] := TraduitGA('Affecté');
  fGD_PlanCharge.CellValues[cInColRealise,0] := TraduitGA('Engagé');
  fGD_PlanCharge.CellValues[cInColRafTh,0] := TraduitGA('RAE Th');
  fGD_PlanCharge.CellValues[cInColRaf,0]   := TraduitGA('RAE');
  fGD_PlanCharge.CellValues[cInColTotal,0] := TraduitGA('Total');
  fGD_PlanCharge.CellValues[cInColEcart,0] := TraduitGA('Ecart');

  // taille des colonnes pour les montants
  fGD_PlanCharge.ColWidths[cInColPrevu] := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColAffecte] := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColRealise] := cInTitre2;

  fGD_PlanCharge.ColWidths[4] := cInMois; fGD_PlanCharge.ColWidths[5] := cInMois;
  fGD_PlanCharge.ColWidths[6] := cInMois; fGD_PlanCharge.ColWidths[7] := cInMois;
  fGD_PlanCharge.ColWidths[8] := cInMois; fGD_PlanCharge.ColWidths[9] := cInMois;
  fGD_PlanCharge.ColWidths[10] := cInMois; fGD_PlanCharge.ColWidths[11] := cInMois;
  fGD_PlanCharge.ColWidths[12] := cInMois; fGD_PlanCharge.ColWidths[13] := cInMois;
  fGD_PlanCharge.ColWidths[14] := cInMois; fGD_PlanCharge.ColWidths[15] := cInMois;

  fGD_PlanCharge.ColWidths[cInColRafTh] := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColRaf]   := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColTotal] := cInTitre2;
  fGD_PlanCharge.ColWidths[cInColEcart] := cInTitre4;

  TFVierge(Ecran).HMTrad.ResizeGridColumns(fGD_PlanCharge);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : sauvegarde d'une ligne de planning
Mots clefs ... :
*****************************************************************}
function TOF_AFPLANCHARGE_MNT.PrepareSauvTob(pTobTaches, pTobLignesPla, pTobLigne : Tob; var pInNumLigne : Integer; TQR : TQRProgressForm) : Boolean;
var
  vTob              : Tob;
  vRdQte            : Double;
  vDtLigne          : TDateTime;
  i                 : Integer;
  vStAff0, vStAff1  : String;
  vStAff2, vStAff3  : String;
  vStAvenant        : String;
  vStTiers          : String;
  vArticle          : RecordArticle;
  vTobTache         : Tob;
  vRdQteAPlanif     : Double;
  vDtFin            : TDateTime;

begin

  Result := True;
  MoveCurProgressForm ('');

  // on se positionne sur la tache
  vTobTache := pTobTaches.FindFirst(['ATA_AFFAIRE', 'ATA_NUMEROTACHE'],[fStAffaire, pTobLigne.GetValue('PC_NUMEROTACHE')], False);

  vRdQteAPlanif := valeur(pTobLigne.GetValue('PC_RESTEAPASSER')) -
                   valeur(pTobLigne.GetValue('PC_PREVU')) +
                   valeur(pTobLigne.GetValue('PC_AFFECTE')) +
                   valeur(pTobLigne.GetValue('PC_REALISE'));

  vTobTache.PutValue('ATA_RAPPTPR',floatToStr(vRdQteAPlanif));

  // valorisation des taches et des ressources
  //Valorisation(vTobTache, 'ATA');

  CodeAffaireDecoupe(fStAffaire, vStAff0, vStAff1, vStAff2, vStAff3,vStAvenant, taModif, false);

  // C.B prise en compte des affaires modeles
  if TeststCleAffaire(fStAffaire,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant,vStTiers, False, True, False, False) <> 1 then
    vStTiers := '';

  GetArticle(pTobLigne.GetValue('PC_ARTICLE'), vArticle);
 
  // parcours de la ligne du plan de charge
  for i := cInColDebMois to cInColFinMois do
    begin
      MoveCurProgressForm ('');
      vDtLigne := CalculDate(i-cInColDebMois);
      vRdQte := TobStrToFloat(pTobLigne.GetValeur(1000 + i));

      // on insere un enregistrement que si qte est non nulle
      if vRdQte <> 0 then
        Begin                 
          vTob := Tob.create('AFPLANNING', pTobLignesPla, -1);
          Try
            vTob.PutValue('APL_AFFAIRE', fStAffaire);
            vTob.PutValue('APL_NUMEROLIGNE', pInNumLigne);
            vTob.PutValue('APL_TYPELIGNEPLA', 'TAC');
            vTob.PutValue('APL_NUMEROTACHE', pTobLigne.GetValue('PC_NUMEROTACHE'));
            vTob.PutValue('APL_RESSOURCE', pTobLigne.GetValue('PC_RESSOURCE'));
            vTob.PutValue('APL_FONCTION', pTobLigne.GetValue('PC_FONCTION'));
            vTob.PutValue('APL_ARTICLE', pTobLigne.GetValue('PC_ARTICLE'));
            vTob.PutValue('APL_TYPEARTICLE', vArticle.StTypeArticle);
            vTob.PutValue('APL_CODEARTICLE', vArticle.StCodeArticle);
            vTob.PutValue('APL_TIERS', vStTiers);
            vTob.PutValue('APL_ETATLIGNE', getparamsoc('SO_AFPLANNINGETAT'));
            vTob.PutValue('APL_TERMINE', '-');
            vTob.PutValue('APL_LIGNEGENEREE', '-');
            vTob.PutValue('APL_ACTIVITEGENERE', '-');
            vTob.PutValue('APL_AFFAIRE0', vStAff0);
            vTob.PutValue('APL_AFFAIRE1', vStAff1);
            vTob.PutValue('APL_AFFAIRE2', vStAff2);
            vTob.PutValue('APL_AFFAIRE3', vStAff3);
            vTob.PutValue('APL_AVENANT', vStAvenant);
            vTob.PutValue('APL_ACTIVITEEFFECT','-');
            vTob.PutValue('APL_UNITETEMPS', getparamsoc('SO_AFMESUREACTIVITE'));
            vTob.PutValue('APL_QTEPLANIFIEE', FloatToStr(0));
            vTob.PutValue('APL_QTEPLANIFUREF', FloatToStr(0));
            vTob.PutValue('APL_QTEREALISE', FloatToStr(0));
            vTob.PutValue('APL_QTEREALUREF', FloatToStr(0));
            vTob.PutValue('APL_INITPTPR', FloatToStr(vRdQte));
            vTob.PutValue('APL_REALPTPR', FloatToStr(0));
            vTob.PutValue('APL_DATEDEBPLA', vDtLigne);

            vDtFin := plusDate(vDtLigne, round(vRdQte - 1), 'J');
            vTob.PutValue('APL_DATEFINPLA', vDtFin);

            vTob.PutValue('APL_LIBELLEPLA', pTobLigne.GetValue('PC_TACHE'));

            // valorisation des lignes de planning
            //Valorisation(vTob, 'APL');

            pInNumLigne := pInNumLigne + 1;
          Except
            PGIBoxAF (TexteMsgTache[3],'');
            result := false;
          End;
        end;
    End;
end;

Initialization
  registerclasses ( [ TOF_AFPLANCHARGE, TOF_AFPLANCHARGE_ANCETRE,
                      TOF_AFPLANCHARGE_QTE, TOF_AFPLANCHARGE_MNT] ) ;
end.


