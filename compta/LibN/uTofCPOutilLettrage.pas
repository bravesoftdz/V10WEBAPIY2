{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPOUTILLETTRAGE ()
Mots clefs ... : TOF;CPOUTILLETTRAGE
*****************************************************************}
Unit uTofCPOutilLettrage;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
{$IFDEF EAGLCLIENT}
     eMul,
     MainEAgl,        // AGLLanceFiche
     UtileAGL,        // PrintDBGrid
{$ELSE}
     Db,
     Fe_Main,         // AGLLanceFiche
     PrintDBG,        // PrintDBGrid
  {$IFDEF DBXPRESS}uDbxDataSet,{$ELSE}Dbtables,{$ENDIF}
     mul,
{$ENDIF}
     Forms,
     Sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu, // THSystemMenu
     Ent1,     // VH^.
     ParamSoc, // GetParamSocSecur
     Htb97,
     uTob,
     HXlspas,      // ExportGrid
     Dialogs,      // TSaveDialog
     uLibEcriture, // CSetMontant
     utilPgi,      // ModePaiementParDefaut
     ED_TOOLS,     // InitMoveProgressForm
     uTobDebug,    // TobDebug(Result);
     uTof ;

Type
  TOF_CPOUTILLETTRAGE = Class (TOF)

    TNbPasEquilibre : THLabel;
    TNbEquilibre    : THLabel;

    FListe          : THGrid;

    E_General       : THEdit;
    E_General_      : THEdit;
    E_Auxiliaire    : THEdit;
    E_Auxiliaire_   : THEdit;
    DateCorrection  : THEdit;
    E_DatePaquetMin : THEdit;
    E_DatePaquetMax : THEdit;
    E_Lettrage      : THEdit;

    E_Devise        : THValComboBox;

    Pages           : TPageControl;
    TabDetail       : TTabSheet;

    BVerif          : TToolBarButton97;
    BDetail         : TToolBarbutton97;
    BCorrection     : TToolBarButton97;
    BExport         : TToolBarButton97;
    BImprimer       : TToolBarButton97;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnCancel                 ; override ;
    procedure OnClose                  ; override ;

    procedure OnClickBVerif      (Sender : TObject);
    procedure OnClickBDetail     (Sender : TObject);
    procedure OnClickBCorrection (Sender : TObject);
    procedure OnClickBExport     (Sender : TObject);
    procedure OnClickBImprimer   (Sender : TObject);
    procedure AuxiElipsisClick   (Sender : TObject);

  private
    FTobFListe : Tob;

    FStWhereCriteres : string;
    FStExoV8         : string;
    FStJal           : string;
    FStCptDebit      : string;
    FStCptCredit     : string;
    FStLibelle       : string;

    FInfoEcriture : TInfoEcriture;
    FMessage      : TMessageCompta;
    FSaveDialog   : TSaveDialog;

    procedure InitFListe;                   
    procedure ChargeFTobFListe;
    procedure RemplitFListe;
    function  RecupereCriteres : string;
    function  ControleParamSocLettrage : Boolean;

    function GenereEcrLettrable(vTobIt : Tob; vNumeroPiece : integer; vStModePaiement, vStModeSaisie : string) : Tob;
    function GenereEcrRegul( vTobIt : Tob; vNumeroPiece : integer; vStModePaiement, vStModeSaisie : string) : Tob;

    procedure CorrectionLettrage;

  end;

procedure CPLanceFiche_CPOutilLettrage;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Constantes,       // ets_Nouveau
  uLibAnalytique,   // AllouAxe
  SaisUtil,         // GetNewNumJal
  VerDPaq,          // DatePaquet
  UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

const cColGene = 1;
      cColAuxi = 2;
      cColCodeLettre = 3;
      cColEcart = 4;
      cColDevise = 5;
      cColLetDevise = 6;
      cColDatePaquetMin = 7;
      cColDatePaquetMax = 8;

      cBigSize    = 100;
      cMiddleSize = 80;
      cLittleSize = 60;

      cStPasEquilibre = 'Nombre de paquets lettrés non équilibrés : ';
      cStEquilibre    = 'Nombre de paquets lettrés corrigés : ';


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CPLanceFiche_CPOutilLettrage;
var lSt1, lSt2, lSt3 : HString;
begin
  // GCO - 02/01/2008 - FQ 21295
  lSt1 := TraduireMemoire('Attention. Il est fortement recommandé de sauvegarder ' +
          'la base avant de lancer le nouvel utilitaire de réparation du lettrage.');

  lSt2 := TraduireMemoire('Voulez-vous continuer ?');

  lSt3 := TraduireMemoire('Outil de correction du lettrage');

  if PgiAsk(lSt1 + #13#10 + lSt2, lSt3 ) = MrYes then
    AGLLanceFiche('CP', 'CPOUTILLETTRAGE', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnArgument (S : String ) ;
begin
  Inherited ;
  FListe         := THGrid(GetControl('FLISTE', True));

  TNbPasEquilibre := THLabel(GetControl('TNBPASEQUILIBRE', True));
  TNbEquilibre    := THLabel(GetControl('TNBEQUILIBRE', True));

  DateCorrection  := THEdit(GetControl('DATECORRECTION', True));
  E_General       := THEdit(GetControl('E_GENERAL', True));
  E_General_      := THEdit(GetControl('E_GENERAL_', True));
  E_Auxiliaire    := THEdit(GetControl('E_AUXILIAIRE', True));
  E_Auxiliaire_   := THEdit(GetControl('E_AUXILIAIRE_', True));
  E_DatePaquetMin := THEdit(GetControl('E_DATEPAQUETMIN', True));
  E_DatePaquetMax := THEdit(GetControl('E_DATEPAQUETMAX', True));
  E_Lettrage      := THEdit(GetControl('E_LETTRAGE', True));

  E_Devise        := THValComboBox(GetControl('E_DEVISE', True));

  Pages          := TPageControl(GetControl('PAGES', True));
  TabDetail      := TTabSheet(GetControl('TABDETAIL', True ));
  BVerif         := TToolBarButton97(GetControl('BVERIF', True));
  BDetail        := TToolBarbutton97(GetControl('BDETAIL', True));
  BCorrection    := TToolBarButton97(GetControl('BCORRECTION', True));
  BExport        := TToolBarButton97(GetControl('BEXPORT', True));
  BImprimer      := TToolBarButton97(GetControl('BIMPRIMER', True));

  // Evénements
  BVerif.OnClick      := OnClickBVerif;
  BDetail.OnClick     := OnClickBDetail;
  BCorrection.OnClick := OnClickBCorrection;
  BExport.OnClick     := OnClickBExport;
  BImprimer.OnClick   := OnClickBImprimer;

  FSaveDialog := TSaveDialog.Create(Ecran);
  FSaveDialog.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
  FSaveDialog.DefaultExt := 'XLS';
  FSaveDialog.FilterIndex := 1;
  FSaveDialog.Options := FSaveDialog.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

  // Init
  E_General.MaxLength     := VH^.Cpta[fbGene].Lg;
  E_General_.MaxLength    := VH^.Cpta[fbGene].Lg;
  E_Auxiliaire.MaxLength  := VH^.Cpta[fbAux].Lg;
  E_Auxiliaire_.MaxLength := VH^.Cpta[fbAux].Lg;

  DateCorrection.ReadOnly := False;
  DateCorrection.Enabled  := False;
  DateCorrection.Text     := DateToStr(VH^.Encours.Deb);
  BCorrection.Enabled     := False;
  BDetail.Enabled         := False;
  TabDetail.TabVisible    := False;
  E_Devise.ItemIndex      := 0;
  E_DatePaquetMin.Text    := StDate1900;
  E_DatePaquetMax.Text    := StDate2099;

  if E_General.CanFocus then E_General.SetFocus;

  InitFListe;

  FTobFListe := Tob.Create('ECRITURE', nil, -1);

  // GCO - 09/01/2007
  FStExoV8 := '';
  if VH^.ExoV8.Code <> '' then
  begin
    FStExoV8 := ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '" ';
  end;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('E_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.InitFListe;
begin
  FListe.ColAligns[cColGene]          := TaCenter;
  FListe.ColWidths[cColGene]          := cBigSize;

  FListe.ColAligns[cColAuxi]          := TaCenter;
  FListe.ColWidths[cColAuxi]          := cBigSize;

  FListe.ColAligns[cColCodeLettre]    := TaCenter;
  FListe.ColWidths[cColCodeLettre]    := cLittleSize;

  FListe.ColAligns[cColEcart]         := TaRightJustify;
  FListe.ColWidths[cColEcart]         := cBigSize;

  FListe.ColAligns[cColDevise]        := TaCenter;
  FListe.ColWidths[cColDevise]        := cLittleSize;

  FListe.ColAligns[cColLetDevise]     := TaCenter;
  FListe.ColWidths[cColLetDevise]     := cLittleSize;

  FListe.ColAligns[cColDatePaquetMin] := TaCenter;
  FListe.ColWidths[cColDatePaquetMin] := cMiddleSize;

  FListe.ColAligns[cColDatePaquetMax] := TaCenter;
  FListe.ColWidths[cColDatePaquetMax] := cMiddleSize;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnLoad ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnUpdate ;
begin
  Inherited ;
  ChargeFTobFListe;
  RemplitFListe;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/08/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.ChargeFTobFListe;
begin
  if not Assigned(FTobFliste) then Exit;
  FTobFListe.LoadDetailFromSQL('SELECT DISTINCT E_AUXILIAIRE, E_GENERAL, E_LETTRAGE, ' +
    'SUM(E_DEBIT-E_CREDIT) ECART, E_DEVISE, ' +
    'E_LETTRAGEDEV, MIN(E_DATEPAQUETMIN) DATEPAQUETMIN, ' +
    'MAX(E_DATEPAQUETMAX) DATEPAQUETMAX, MAX(E_ETABLISSEMENT) ETAB FROM ECRITURE WHERE ' +
    'E_QUALIFPIECE = "N" AND E_ECHE = "X" AND E_LETTRAGE <> '' AND ' +
    'E_ETATLETTRAGE = "TL" AND E_ECRANOUVEAU <> "OAN" AND ' +
    'E_ECRANOUVEAU <> "C" AND E_LETTRAGEDEV = "-" ' + FStWhereCriteres + FStExoV8 +
    ' GROUP BY E_AUXILIAIRE, E_GENERAL, E_LETTRAGE, E_DEVISE , E_LETTRAGEDEV ' +
    'HAVING SUM(E_DEBIT-E_CREDIT) <> 0 ORDER BY E_GENERAL, E_AUXILIAIRE');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.RemplitFListe;
begin
  if not Assigned(FTobFliste) then Exit;

  FListe.RowCount := 2;

  if FTobFliste.Detail.Count > 0 then
    FListe.RowCount := FTobFListe.Detail.Count + 1;

  FTobFListe.PutGridDetail(FListe, False, False, 'E_GENERAL;E_AUXILIAIRE;E_LETTRAGE;ECART;E_DEVISE;E_LETTRAGEDEV;DATEPAQUETMIN;DATEPAQUETMAX');

  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);

  if FListe.CanFocus then FListe.SetFocus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPOUTILLETTRAGE.RecupereCriteres: string;
begin
  Result := '';

  if E_General.Text <> '' then
    Result := Result + ' AND E_GENERAL >= "'+ E_General.Text + '"';

  if E_General_.Text <> '' then
    Result := Result + ' AND E_GENERAL <= "'+ E_General_.Text + '"';

  if E_Auxiliaire.Text <> '' then
    Result := Result + ' AND E_AUXILIAIRE >= "'+ E_Auxiliaire.Text + '"';

  if E_Auxiliaire_.Text <> '' then
    Result := Result + ' AND E_AUXILIAIRE <= "'+ E_Auxiliaire_.Text + '"';

  if E_Devise.ItemIndex <> 0 then
    Result := Result + ' AND E_DEVISE = "' + E_Devise.Value + '"';

  if E_Lettrage.Text <> '' then
    Result := Result + ' AND E_LETTRAGE = "' + E_Lettrage.Text + '"';

  if E_DatePaquetMin.Text <> StDate1900 then
    Result := Result + ' AND E_DATEPAQUETMIN >= "' + UsDateTime(StrToDate(E_DatePaquetMin.Text)) + '"';

  if E_DatePaquetMax.Text <> StDate2099 then
    Result := Result + ' AND E_DATEPAQUETMAX <= "' + UsDateTime(StrToDate(E_DatePaquetMax.Text)) + '"';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/08/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPOUTILLETTRAGE.ControleParamSocLettrage: Boolean;
begin
  Result := False;

  FStJal := GetParamSocSecur('SO_LETCHOIXJAL', '');
  if Trim(FStJal) = '' then
  begin
    PgIInfo('Vous devez renseigner le journal des écritures de régularisation.', 'Paramètre société - Lettrage Gestion');
    Exit;
  end;

  FStCptDebit := GetParamSocSecur('SO_LETCHOIXGEN', '');
  if Trim(FStCptDebit) = '' then
  begin
    PgIInfo('Vous devez renseigner le compte général débiteur des écritures de régularisation.', 'Paramètre société - Lettrage Gestion');
    Exit;
  end;

  FStCptCredit := GetParamSocSecur('SO_LETCHOIXGENC', '');
  if Trim(FStCptDebit) = '' then
  begin
    PgiInfo('Vous devez renseigner le compte général créditeur des écritures de régularisation.', 'Paramètre société - Lettrage Gestion');
    Exit;
  end;

  FStLibelle   := GetParamSocSecur('SO_LETCHOIXLIB', '');
  if Trim(FStLibelle) = '' then
  begin
    PgiInfo('Vous devez renseigner le libellé des écritures de régularisation.', 'Paramètre société - Lettrage Gestion');
    Exit;
  end;

  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnClickBVerif(Sender: TObject);
var lQuery : TQuery;
    lStSql : string;
begin
  TabDetail.TabVisible := False;
  TNbEquilibre.Caption := cStEquilibre + IntToStr(0);

  try
    FStWhereCriteres := RecupereCriteres;

    lStSql := 'SELECT DISTINCT E_AUXILIAIRE, E_GENERAL, E_LETTRAGE, ' +
      'SUM(E_DEBIT-E_CREDIT) ECART, E_DEVISE, E_LETTRAGEDEV ' +
      'FROM ECRITURE WHERE E_QUALIFPIECE = "N" AND ' +
      'E_ECHE = "X" AND E_LETTRAGE <> '' AND E_ETATLETTRAGE = "TL" AND ' +
      'E_ECRANOUVEAU <> "OAN" AND E_ECRANOUVEAU <> "C" AND E_LETTRAGEDEV = "-" ' +
      FStWhereCriteres + FStExoV8 +
      ' GROUP BY E_AUXILIAIRE, E_GENERAL, E_LETTRAGE, E_DEVISE , ' +
      'E_LETTRAGEDEV HAVING SUM(E_DEBIT-E_CREDIT) <> 0 ORDER BY E_GENERAL, E_AUXILIAIRE';

    lQuery := OpenSql(lStSql, True);
    if not lQuery.Eof then
    begin
      TNbPasEquilibre.Caption := cStPasEquilibre + IntToStr(lQuery.RecordCount);
      TNbPasEquilibre.Font.Color := ClRed;
    end
    else
    begin
      TNbPasEquilibre.Caption := cStPasEquilibre + IntToStr(0);
      TNbPasEquilibre.Font.Color := ClBlack;
      TabDetail.TabVisible := False;
      FTobFListe.ClearDetail;
      FListe.VidePile(False);
      PgiInfo('Vérification terminée. Aucune anomalie détectée.', Ecran.Caption);
    end;

    BDetail.Enabled        := (not lQuery.Eof);
    BCorrection.Enabled    := (not lQuery.Eof);
    DateCorrection.Enabled := (not lQuery.Eof);
  finally
    Ferme(lQuery);
  end;

  // GCO - 09/11/2007 - FQ 21820 + FQ 21295
  if BDetail.Enabled then
  begin
    if PgiAsk('Attention. Suite aux erreurs détectées, il est fortement recommandé d''utiliser ' +
              'l''utilitaire de correction' + #13#10 + ' de "Date des paquets lettrés" pour plus de sécurité. ' +
              'Voulez-vous le lancer maintenant ?',  Ecran.Caption) = MrYes then
      DatePaquet;
  end
  // FIN GCO
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnClickBDetail(Sender: TObject);
begin
  TabDetail.TabVisible := True;
  Pages.ActivePage := TabDetail;
  ChargeFTobFListe;
  RemplitFListe;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnClickBCorrection(Sender: TObject);
var lDtMin, lDtMax : TDateTime;
begin
  lDtMin := VH^.EnCours.Deb;

  if (VH^.Suivant.Code <> '') and (VH^.Suivant.EtatCpta = 'OUV') then
    lDtMax := VH^.Suivant.Fin
  else
    lDtMax := VH^.EnCours.Fin;

  if not ((StrToDate(DateCorrection.Text) >= lDtMin) and
          (StrToDate(DateCorrection.Text) <= lDtMax)) then
  begin
    PgIInfo('La date de correction du lettrage doit être sur un exercice ouvert.', Ecran.Caption);
    Exit;
  end;

  if not ControleParamSocLettrage then Exit;

  if not Assigned(FTobFliste) then Exit;

  // GCO - 03/04/2007 - FQ 19547
  if PgiAsk('Avant d''activer le bouton "Lancer la correction", il vous appartient de consulter ' + #13#10 +
            'le rapport d''erreurs par le bouton "voir le détail" et d''effectuer manuellement les ' + #13#10 +
            'corrections de lettrage. ' + #13#10 +
	          'Dans le cas contraire, le traitement va générer des mouvements afin d''équilibrer ' + #13#10 +
            'le lettrage (Modification soldes comptes). Voulez vous continuer ?',  Ecran.Caption) <> MrYes then
    Exit;

  if FTobFListe.Detail.Count = 0 then
    ChargeFTobFListe;

  if Transactions(CorrectionLettrage, 1) = oeOk then
  begin
    TNbEquilibre.Caption := cStEquilibre + IntToStr(FTobFListe.Detail.Count);
    BDetail.Enabled      := False;
    BCorrection.Enabled  := False;
    PgiInfo('Traitement terminé. Aucune anomalie détectée.', Ecran.Caption);
  end
  else
    PgiInfo('Traitement annulé. Erreur pendant la correction du lettrage.', Ecran.Caption);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnClickBExport(Sender: TObject);
var lStHint: string;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then Exit;
  if FSaveDialog.Execute then
  begin
    if FSaveDialog.FilterIndex = 5 then //html
    begin
      lStHint := FListe.Hint;
      FListe.Hint := Ecran.Caption;
      ExportGrid(FListe, nil, FSaveDialog.FileName, FSaveDialog.FilterIndex, True);
      FListe.Hint := lStHint;
    end
    else
      ExportGrid(FListe, nil, FSaveDialog.FileName, FSaveDialog.FilterIndex, True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnClickBImprimer(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
   PgiInfo('Etat non disponible en CWAS', 'Erreur');
{$ELSE}
   PrintDBGrid(FListe, nil, Ecran.Caption, '', '');
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/08/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.CorrectionLettrage;
var i : integer;
    lTobEcrInsert : Tob;
    lTobEcrCpt    : Tob;
    lTobEcrREgul  : Tob;
    lTobIt        : Tob;

    lInNumeroPiece  : integer;
    lStModePaiement : string;
    lStModeRegl     : string;
    lStModeSaisie   : string;

    // Pour utilisation du Noyau
    lRecError       : TRecError;
begin
  lInNumeroPiece := GetNewNumJal(FStJal, True, StrToDate(DateCorrection.Text), '', '', '');
  if lInNumeroPiece <= 0 then
  begin
    PgiInfo('Erreur lors du calcul du numéro de pièce.', Ecran.Caption);
    Exit;
  end;

  FMessage      := TMessageCompta.Create('');
  FInfoEcriture := TInfoEcriture.Create;
  FInfoEcriture.Aux.BoCodeOnly := True;

  lStModePaiement := ModePaiementParDefaut(lStModeRegl);

  lStModeSaisie := GetColonneSql('JOURNAL', 'J_MODESAISIE', 'J_JOURNAL = "'+ FStJal + '"');

  lTobEcrInsert := nil;
  try
    InitMoveProgressForm( nil, 'Outil de correction du lettrage', 'Veuillez patienter...',  FTobFListe.Detail.Count-1, False, True);
    lTobEcrInsert := Tob.Create('', nil, -1);

    for i := 0 to FTobFListe.Detail.Count -1 do
    begin
      lTobIt := FTobFListe.Detail[i];

      lTobEcrCpt := GenereEcrLettrable( lTobIt, lInNumeroPiece, lStModePaiement, lStModeSaisie );
      lRecError := CIsValidLigneSaisie(lTobEcrCpt, FInfoEcriture);
      if ( lRecError.RC_Error = RC_DATEINCORRECTE ) then
        lRecError.RC_Error := RC_PASERREUR ;

      if (lRecError.RC_Error <> RC_PASERREUR) then
      begin
        PgIInfo('Ecriture lettrable : ' + FMessage.GetMessage(lRecError.RC_Error) + #13#10 + lRecError.RC_Message, Ecran.Caption);
        Break;
      end;

      lTobEcrRegul := GenereEcrRegul( lTobIt, lInNumeroPiece, lStModePaiement, lStModeSaisie );
      lRecError := cIsValidLigneSaisie(lTobEcrRegul, FInfoEcriture);
      if ( lRecError.RC_Error = RC_DATEINCORRECTE ) then
        lRecError.RC_Error := RC_PASERREUR ;

      if (lRecError.Rc_Error <> RC_PASERREUR) then
      begin
        PgiInfo('Ecriture de régularisation : ' + FMessage.GetMessage(lRecError.RC_Error) + #13#10 + lRecError.RC_Message, Ecran.Caption);
        Break;
      end;

      lTobEcrCpt.ChangeParent( lTobEcrInsert, -1 );
      lTobEcrRegul.ChangeParent( lTobEcrInsert, -1 );

      // Affectation de la contrepartie
      CAffectCompteContrePartie( lTobEcrInsert );

      // Vérification de la pièce
      if not CEnregistreSaisie( lTobEcrInsert, False, False , True , FInfoEcriture ) then
        PgiInfo('Fonction d''appel CEnregistreSaisie : ' + V_Pgi.LastSQLError, Ecran.Caption)
      else
      begin
        // GCO - CORRECTION BUG ISVALIDAUX DU NOYAU QUI MET #0
        if lTobIt.GetString('E_AUXILIAIRE') = '' then
          lTobEcrCpt.SetString('E_AUXILIAIRE', '');

        // Mise à jour des dates paquets MIN et MAX
        ExecuteSQL('UPDATE ECRITURE SET ' +
          'E_DATEPAQUETMIN = "' + UsDateTime(lTobEcrCpt.GetDateTime('E_DATEPAQUETMIN')) + '", ' +
          'E_DATEPAQUETMAX = "' + UsDateTime(lTobEcrCpt.GetDateTime('E_DATEPAQUETMAX')) + '", ' +
          'E_DATEMODIF = "' + UsTime(Now) + '" WHERE ' +
          'E_GENERAL = "' + lTobEcrCpt.GetString('E_GENERAL') + '" AND ' +
          'E_AUXILIAIRE = "' + lTobEcrCpt.GetString('E_AUXILIAIRE') + '" AND ' +
          'E_LETTRAGE = "' + lTobEcrCpt.GetString('E_LETTRAGE') + '"');

        Inc( lInNumeroPiece );

        lTobEcrCpt.Free;
        lTobEcrRegul.Free;
      end;
      MoveCurProgressForm('Traitement du lettrage en cours...');
    end;

  finally
    FiniMoveProgressForm;
    lTobEcrInsert.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPOUTILLETTRAGE.OnClose ;
begin
  FTobFListe.Free;
  FreeAndNil(FInfoEcriture);
  FreeAndNil(FMessage);
  FreeAndNil(FSaveDialog);
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function TOF_CPOUTILLETTRAGE.GenereEcrLettrable(vTobIt: Tob; vNumeroPiece : integer; vStModePaiement, vStModeSaisie : string) : Tob;
var lDtMin : TDateTime;
    lDtMax : TDateTime;
    lDbEcart : Double;
    lDevise : RDevise;
begin
  if StrToDate(DateCorrection.Text) > vTobIt.GetDateTime('DATEPAQUETMIN') then
    lDtMin := vTobIt.GetDateTime('DATEPAQUETMIN')
  else
    lDtMin := StrToDate(DateCorrection.Text);

  if StrToDate(DateCorrection.Text) > vTobIt.GetDateTime('DATEPAQUETMAX') then
    lDtMax := StrToDate(DateCorrection.Text)
  else
    lDtMax := vTobIt.GetDateTime('DATEPAQUETMAX');

  lDbEcart := vTobIt.GetDouble('ECART');

  lDevise.Code := vTobIt.GetString('E_DEVISE');
  GetInfosDevise(lDevise);
  lDevise.Taux := GetTaux(lDevise.Code , lDevise.DateTaux, V_PGI.DateEntree) ;

  Result := Tob.Create('ECRITURE', nil, -1);
  CPutDefautEcr(Result);

  CRemplirDateComptable( Result, StrToDate(DateCorrection.Text));
  Result.SetDateTime('E_DATECREATION', Date);
  Result.SetDateTime('E_DATEMODIF', Now);

  // Init des champs de Lettrage
  CRemplirInfoLettrage( Result );
  Result.SetString('E_ETATLETTRAGE', 'TL');
  Result.SetDouble('E_COUVERTURE', Abs(lDbEcart));
  Result.SetString('E_LETTRAGE', vTobIt.GetString('E_LETTRAGE'));
  Result.SetString('E_LETTRAGEDEV', vTobIt.GetString('E_LETTRAGEDEV'));
  Result.SetDateTime('E_DATEPAQUETMIN', lDtMin);
  Result.SetDateTime('E_DATEPAQUETMAX', lDtMax);
  Result.SetDouble('E_COUVERTUREDEV', Abs(lDbEcart));

  Result.SetString('E_GENERAL', vTobIt.GetString('E_GENERAL'));

  if vTobIt.GetString('E_AUXILIAIRE') <> '' then
    Result.SetString('E_AUXILIAIRE', vTobIt.GetString('E_AUXILIAIRE'))
  else
    Result.SetString('E_AUXILIAIRE', '');

  Result.SetString('E_EXERCICE', QUELEXODT(StrToDateTime(DateCorrection.Text)));
  Result.SetString('E_JOURNAL', FStJal);
  Result.SetString('E_ETABLISSEMENT', vTobIt.GetString('ETAB'));
  Result.SetString('E_DEVISE', lDevise.Code);
  Result.SetString('E_IO', 'X');
  Result.SetString('E_LIBELLE', TraduireMemoire('Outil de correction du lettrage'));
  Result.SetInteger('E_NUMEROPIECE', vNumeroPiece);
  Result.SetInteger('E_NUMLIGNE', 1);
  Result.SetString('E_MODEPAIE', vStModePaiement);
  Result.SetString('E_MODESAISIE', vStModeSaisie);
  Result.SetDateTime('E_DATEECHEANCE', StrToDate(DateCorrection.Text));

  // GCO - 08/11/2007 - FQ 20537

  if lDbEcart > 0 then
    CSetMontants(Result, 0, lDbEcart, lDevise, True)
  else
    CSetMontants(Result, Abs(lDbEcart), 0, lDevise, True);

  CGetEch( Result, FInfoEcriture );
  CGetRegimeTva( Result, FInfoEcriture );

  FInfoEcriture.LoadCompte( Result.GetString('E_GENERAL'));

  if FInfoEcriture.Compte.IsVentilable then
  begin
    AlloueAxe( Result );
    CVentilerTOB( Result, FInfoEcriture );
    Result.SetString('E_ANA', 'X');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPOUTILLETTRAGE.GenereEcrRegul(vTobIt: Tob; vNumeroPiece : integer; vStModePaiement, vStModeSaisie : string) : Tob;
var lDbEcart : Double;
    lStGeneral : string;
    lDevise : RDevise;
begin
  lDevise.Code := vTobIt.GetString('E_DEVISE');
  GetInfosDevise(lDevise);
  lDevise.Taux := GetTaux(lDevise.Code , lDevise.DateTaux, V_PGI.DateEntree) ;

  Result := Tob.Create('ECRITURE', nil, -1);
  CPutDefautEcr( Result );

  CRemplirDateComptable( Result, StrToDate(DateCorrection.Text));
  Result.SetDateTime('E_DATECREATION', Date);
  Result.SetDateTime('E_DATEMODIF', Now);

  CSupprimerInfoLettrage( Result );
  lDbEcart := vTobIt.GetDouble('ECART');
  if lDbEcart > 0 then
    lStGeneral := FStCptDebit
  else
    lStGeneral := FStCptCredit;

  FInfoEcriture.LoadCompte(lStGeneral);
  if FInfoEcriture.Compte.IsLettrable then
    CRemplirInfoLettrage( Result )
  else
    if FInfoEcriture.Compte.IsPointable then
      CRemplirInfoPointage( Result );

  Result.SetString('E_GENERAL', lStGeneral);
  Result.SetString('E_AUXILIAIRE', '');
  Result.SetString('E_EXERCICE', QUELEXODT(StrToDateTime(DateCorrection.Text)));
  Result.SetString('E_JOURNAL', FStJal);
  Result.SetString('E_ETABLISSEMENT', vTobIt.GetString('ETAB'));
  Result.SetString('E_DEVISE', lDevise.Code);
  Result.SetString('E_IO', 'X');
  Result.SetString('E_TRESOSYNCHRO', 'CRE');
  Result.SetString('E_LIBELLE', FStLibelle);
  Result.SetInteger('E_NUMEROPIECE', vNumeroPiece);
  Result.SetInteger('E_NUMLIGNE', 2);
  Result.SetString('E_MODEPAIE', vStModePaiement);
  Result.SetString('E_MODESAISIE', vStModeSaisie);
  Result.SetString('E_IO', 'X');

  // GCO - 08/11/2007 - FQ 20537

  if lDbEcart > 0 then
    CSetMontants(Result, lDbEcart, 0, lDevise, True)
  else
    CSetMontants(Result, 0, Abs(lDbEcart), lDevise, True);

  CGetEch( Result, FInfoEcriture );
  CGetRegimeTva( Result, FInfoEcriture );

  if FInfoEcriture.Compte.IsVentilable then
  begin
    AlloueAxe( Result );
    CVentilerTOB( Result, FInfoEcriture );
    Result.SetString('E_ANA', 'X');
  end;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPOUTILLETTRAGE ] ) ;
end.
