{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/03/2004
Modifié le ... : 10/08/2004
Description .. : Source TOF de la FICHE : AMINTEG ()
Suite ........ : - CA - 10/08/2004 : FQ 13797 : suppression de la possibilité
Suite ........ : de gérer les abonnements. Cette fonction était présente
Suite ........ : mais non géreé complètement (pas de création de guide
Suite ........ : associé au contrat d'abonnement)
Suite ........ :  PGR - 27/01/2006 - Nouveaux paramètres société pour intégration des écritures
Suite......... : MBO - 02/06/2006 - FQ 15561 - disable double mouette si exercice <> compta
Suite......... : MVG - 12/07/2006 - Correction pour conseil compilation SERIE1
Suite......... : MBO - 21/11/2006 - FQ 19070 - exclure de la liste les immos financières et tous les biens SIMU
Mots clefs ... : TOF;AMINTEG
*****************************************************************}
Unit AMINTEG_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,                                            
     Windows,
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
     HDB,
//     aboutil,
//     Contabon,
//     ImCrGuid,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     paramsoc,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     AGLInit,
     HStatus,
     LookUp,
     ImEnt,
     IMMO_TOM,
     ImContra,
     IntegEcr,
     ImOutGen ;

Type
  TOF_AMINTEG = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      {$IFDEF EAGLCLIENT}
      fListe : THGrid;
      {$ELSE}
      fListe : THDBGrid;
      {$ENDIF}
      fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:
        TShiftState) of object;
      procedure OnChangeTypeEcriture ( Sender : TObject );
      procedure OnChangeChampLibre ( Sender : TObject );
      procedure OnElipsisCompte ( Sender : TObject );
      procedure OnListeDblClick ( Sender : TObject );
//      procedure OnAbonnementClick ( Sender : TObject );
      procedure OnSelectAllClick ( Sender : TObject );
      procedure OnOuvrirClick ( Sender : TObject );
      procedure   OnFormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure RecupereListeImmo( L : HTStringList ) ;
{$IFDEF SERIE1}
{$ELSE}
//      procedure InitCommunContrat (Code,Libelle : string; OB : TOB);
{$ENDIF}
  end ;

procedure AMLanceFiche_ListeDesIntegrations ( TypeInt : string = '' );
//function CalculPeriodeEqAbo (Periode : string) : string;

Implementation

const

//PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
 HM: array[0..0] of string =
 {1} ('Veuillez définir les paramètres d''intégration d''écritures dans la commande paramètres.'
    );

procedure AMLanceFiche_ListeDesIntegrations ( TypeInt : string = '' );
var stAction : string;
begin
  stAction := TypeInt;

  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  {$IFNDEF SERIE1}
  if (GetParamSocSecur('SO_IMMOJALDOTDEF','') = '') or (GetParamSocSecur('SO_IMMODOTCHOIXDET','') = '') or
     (GetParamSocSecur('SO_IMMODOTCHOIXTYP','') = '') then
  begin
    PGIBox(HM[0]);
    exit;
  end;
  {$ENDIF}

  AGLLanceFiche('AM','AMINTEG','','',stAction);
end;

procedure TOF_AMINTEG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMINTEG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMINTEG.OnUpdate ;
begin
  inherited;
end ;

procedure TOF_AMINTEG.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMINTEG.OnArgument (S : String ) ;
var okok : boolean;
begin
  Inherited ;
  if S <> '' then SetControlText('TYPEECRITURE',S);
  SetControlText('I_NATUREIMMO','');
  {$IFDEF SERIE1}
  if (GetControlText('TYPEECRITURE')='') then SetControlText('TYPEECRITURE','DOT')
  else SetControlEnabled('TYPEECRITURE',False);
  {$ELSE}
  if (GetControlText('TYPEECRITURE')='') then SetControlText('TYPEECRITURE','DOT');
  {$ENDIF}
  OnChangeTypeEcriture(nil) ;
  {$IFDEF SERIE1}
  if (GetControlText('TYPEECRITURE')='DOT') then Ecran.HelpContext:=541000 else Ecran.HelpContext:=541100 ;
  {$ELSE}
  Ecran.HelpContext:=2410000 ;
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  fListe := THGrid (GetControl ('FLISTE'));
  {$ELSE}
  fListe := THDBGrid (GetControl ('FLISTE'));
  {$ENDIF}
  fListe.OnDblClick := OnListeDblClick;
  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  ECRAN.OnKeyDown := OnFormKeyDown;
  THValComboBox(GetControl('TYPEECRITURE')).OnChange := OnChangeTypeEcriture;
  THCritMaskEdit(GetControl('I_COMPTEIMMO')).OnElipsisClick := OnElipsisCompte;
  THCritMaskEdit(GetControl('I_COMPTELIE')).OnElipsisClick := OnElipsisCompte;
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;
//  TToolbarButton97(GetControl('BABONNEMENT')).OnClick := OnAbonnementClick;
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := OnSelectAllClick;
  TToolbarButton97(GetControl('BOUVRIR')).OnClick := OnOuvrirClick;
  OnChangeTypeEcriture(nil);
  {$IFDEF SERIE1}
  ImLibellesTableLibre(TTabSheet(GetControl('PzLibreS1')),'TT_TABLELIBREIMMO','','I') ;
  Okok:=false ;
  {$ELSE}
  ImLibellesTableLibre(TTabSheet(GetControl('PzLibre')),'TI_TABLE','I_TABLE','I') ;
  Okok:=true ;
  {$ENDIF}
  SetControlVisible('PzLibre',Okok );
  SetControlVisible('PzLibreS1',not Okok);
  SetControlVisible('BABONNEMENT',False); // FQ 13797
  // ajout mbo FQ 15561 - modif le 02.06.2006
  if (GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Encours.Code) then
     SetControlEnabled('bOuvrir', False);
end ;

procedure TOF_AMINTEG.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMINTEG.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMINTEG.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AMINTEG.OnChangeTypeEcriture(Sender: TObject);
begin
//  if (fListe <> nil) then SetControlEnabled('BABONNEMENT', (GetControlText('TYPEECRITURE') = 'ECH'));

  {modif mbo - fq 19070 - 21/11/2006 - on ne prend plus en compte les immos financières
                             ni tous les biens de qualif SIMU
  if (GetControlText('TYPEECRITURE') = 'DOT') then
    SetControlText('XX_WHERE','I_QUALIFIMMO<>"REG" AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_ETAT<>"FER"')
  else if (GetControlText('TYPEECRITURE') = 'ECH') then
    SetControlText('XX_WHERE','I_QUALIFIMMO<>"REG" AND (I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND I_ETAT<>"FER"')
  else SetControlText('XX_WHERE','I_QUALIFIMMO<>"REG" AND I_ETAT<>"FER"');
  }

  if (GetControlText('TYPEECRITURE') = 'DOT') then
    SetControlText('XX_WHERE','((I_QUALIFIMMO<>"REG") AND (I_QUALIFIMMO<>"S")) AND (I_NATUREIMMO="PRO" ) AND I_ETAT<>"FER"')
  else if (GetControlText('TYPEECRITURE') = 'ECH') then
    SetControlText('XX_WHERE','((I_QUALIFIMMO<>"REG") AND (I_QUALIFIMMO<> "S")) AND (I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND I_ETAT<>"FER"')
  else SetControlText('XX_WHERE','((I_QUALIFIMMO<>"REG") AND (I_QUALIFIMMO<> "S")) AND I_ETAT<>"FER"');
end;

procedure TOF_AMINTEG.OnElipsisCompte(Sender: TObject);
var stWhere : string;
begin
  if THCritMaskEdit(Sender).Name = 'I_COMPTEIMMO' then
  stWhere := '(G_GENERAL>="'+VHImmo^.CpteImmoInf+'" AND G_GENERAL<="'+VHImmo^.CpteImmoSup+'") OR '+
             '(G_GENERAL>="'+VHImmo^.CpteFinInf+'" AND G_GENERAL<="'+VHImmo^.CpteFinSup+'")'
  else if THCritMaskEdit(Sender).Name = 'I_COMPTELIE' then
    stWhere := '(G_GENERAL>="'+VHImmo^.CpteCBInf+'" AND G_GENERAL<="'+VHImmo^.CpteCBSup+'") OR '+
               '(G_GENERAL>="'+VHImmo^.CpteLocInf+'" AND G_GENERAL<="'+VHImmo^.CpteLocSup+'")';
  LookupList(TControl(Sender),'','GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,0)  ;
end;

procedure TOF_AMINTEG.RecupereListeImmo(L: HTStringList);
var i : integer;
begin
  if L= nil then exit ;
  {$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
  {$ELSE}
  TheMulQ := TFMul(Ecran).Q;
  {$ENDIF}
  if FListe.AllSelected then
  begin
    InitMove(TheMulQ.RecordCount,'');
    TheMulQ.First;
    while not TheMulQ.EOF do
    begin
      MoveCur(False);
      L.Add(TheMulQ.FindField('I_IMMO').AsString);
      TheMulQ.Next;
    end;
    FListe.AllSelected:=False;
  end else
  begin
    InitMove(FListe.NbSelected,'');
    for i:=0 to FListe.NbSelected-1 do
    begin
      MoveCur(False);
      FListe.GotoLeBookmark(i);
      L.Add(GetDataSet.FindField('I_IMMO').AsString );
    end;
    FListe.ClearSelected;
  end;
  FiniMove;
end;

procedure TOF_AMINTEG.OnListeDblClick(Sender: TObject);
begin
  AMLanceFiche_FicheImmobilisation( GetField('I_IMMO'), taConsult ,'') ;
end;
(*
procedure TOF_AMINTEG.OnAbonnementClick(Sender: TObject);
{$IFDEF SERIE1}
begin
{$ELSE}
{$IFDEF CCS3}
BEGIN
{$ELSE}
var OB, OBAbo : TOB;
    QImmo, QContrat : TQuery;
    Contrat : TImContrat;
    stCodeAbo, stPeriode : string;
    i : integer;
    ARecord : PTranche;
begin
  inherited;
  stCodeAbo := GetField('I_NATUREIMMO') + GetField('I_IMMO') + '01';
  QContrat := OpenSQL ('SELECT * FROM CONTABON WHERE CB_CONTRAT="'+stCodeAbo+'"',True);
{$IFDEF EAGLCLIENT}
{$ELSE}
  if not QContrat.Eof then ParamAbonnement(True,stCodeAbo,taConsult)
  else
  begin
    QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+GetField('I_IMMO')+'"',True);
    stPeriode := QImmo.FindField('I_PERIODICITE').AsString;
    Contrat := TImContrat.Create;
    Contrat.Charge (QImmo);
    Contrat.ChargeTableEcheance;
    Ferme (QImmo);
    OB := TOB.Create ('',nil,-1);
    Contrat.ConvertEcheanceIntoTranche;
    if Contrat.sTypeLoyer = 'LCO' then
    begin // Cas loyer constant
      OBAbo := TOB.Create ('CONTABON',OB,-1);
      InitCommunContrat (stCodeAbo,GetField('I_LIBELLE'),OBAbo);
      ARecord := Contrat.ListeTranches.Items[0];
      OBAbo.PutValue('CB_DATECONTRAT',ARecord^.DateDebut);
      OBAbo.PutValue('CB_SEPAREPAR',CalculPeriodeEqAbo (stPeriode));
      OBAbo.PutValue('CB_NBREPETITION',ARecord^.nEcheance);
      CreationGuideEcheance (OBAbo);
    end
    else
    begin  // Cas  loyer variable
      for i := 1 to Contrat.ListeTranches.Count do
      begin
        OBAbo := TOB.Create ('CONTABON',OB,-1);
        stCodeAbo := GetField('I_NATUREIMMO') + Contrat.sCode + Format('%.2d',[i]);
        ARecord := Contrat.ListeTranches.Items[i-1];
        InitCommunContrat (stCodeAbo,GetField('I_LIBELLE'),OBAbo);
        OBAbo.PutValue('CB_DATECONTRAT',ARecord^.DateDebut);
        OBAbo.PutValue('CB_SEPAREPAR',CalculPeriodeEqAbo (stPeriode));
        OBAbo.PutValue('CB_NBREPETITION',ARecord^.nEcheance);
        CreationGuideEcheance (OBAbo);
      end;
    end;
    OB.InsertDB(nil);
    Contrat.Free;
    OB.Free;
  end;
{$ENDIF}
  Ferme (QContrat);
{$ENDIF}
{$ENDIF}
end;
*)

procedure TOF_AMINTEG.OnSelectAllClick(Sender: TObject);
begin
  FListe.AllSelected := not FListe.AllSelected;
//  SetControlEnabled('BABONNEMENT',(GetControlText('TYPEECRITURE') = 'ECH'));
end;

procedure TOF_AMINTEG.OnChangeChampLibre( Sender : TObject );
begin
  if THValComboBox(Sender).Name='TABLELIBRE1' then SetControlText('I_TABLE0',GetCOntrolText('TABLELIBRE1'))
  else if THValComboBox(Sender).Name='TABLELIBRE2' then SetControlText('I_TABLE2',GetCOntrolText('TABLELIBRE2'))
  else if THValComboBox(Sender).Name='TABLELIBRE3' then SetControlText('I_TABLE3',GetCOntrolText('TABLELIBRE3'));
end;

{$IFDEF SERIE1}
{$ELSE}
(*
procedure TOF_AMINTEG.InitCommunContrat (Code,Libelle : string; OB : TOB);
begin
  OB.PutValue('CB_CONTRAT',Code);
  OB.PutValue('CB_COMPTABLE','X');
  OB.PutValue('CB_LIBELLE',Libelle);
  OB.PutValue('CB_ARRONDI','PAS');
  OB.PutValue('CB_RECONDUCTION','SUP');
  OB.PutValue('CB_DEJAGENERE',0);
  OB.PutValue('CB_DATEDERNGENERE',iDate1900);
  OB.PutValue('CB_GUIDE','');
  OB.PutValue('CB_DATECREATION',Date);
  OB.PutValue('CB_DATEMODIF',Date);
  OB.PutValue('CB_UTILISATEUR',V_PGI.User);
  OB.PutValue('CB_SOCIETE',V_PGI.CodeSociete);
end;
*)
{$ENDIF}

procedure TOF_AMINTEG.OnOuvrirClick ( Sender : TObject );
var ListeImmo : HTStringList;
    Nature : string;
begin
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  {$IFNDEF SERIE1}
  if (GetControlText('TYPEECRITURE') = 'DOT') then
  begin
    if (GetParamSocSecur('SO_IMMOJALDOTDEF','') = '') or (GetParamSocSecur('SO_IMMODOTCHOIXDET','') = '') or
       (GetParamSocSecur('SO_IMMODOTCHOIXTYP','') = '') then
    begin
      PGIBox(HM[0]);
      exit;
    end;
  end
  else
  begin
    if (GetParamSocSecur('SO_IMMOJALECHDEF','') = '') or (GetParamSocSecur('SO_IMMOECHCHOIXDET','') = '') or
       (GetParamSocSecur('SO_IMMOECHCHOIXTYP','') = '') then
    begin
      PGIBox(HM[0]);
      exit;
    end;
  end;
  {$ENDIF}

  ListeImmo:=HTStringList.Create ;
  try
    RecupereListeImmo(ListeImmo) ;
    if ListeImmo.Count <> 0 then
    begin
      Nature := GetField('I_NATUREIMMO');
      if (Nature = 'PRO') or (Nature = 'FI') then
        IntegrationEcritures (toDotation,ListeImmo,TRUE,FALSE)
      else if (Nature = 'CB') or (Nature = 'LOC') then
        IntegrationEcritures (toEcheance,ListeImmo,TRUE,FALSE);
    end;
    FListe.ClearSelected;
  finally
    ListeImmo.Free ;
  end ;
end;

(*
function CalculPeriodeEqAbo (Periode : string) : string;
begin
  if Periode = 'MEN' then Result := '1M'
  else if Periode = 'TRI' then Result := '3M'
  else if Periode = 'SEM' then Result := '6M'
  else if Periode = 'ANN' then Result := '12M'
  else Result := '1M';
end;
*)

procedure TOF_AMINTEG.OnFormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F10 :
      begin
        TToolBarButton97(GetControl('BOUVRIR')).Click;
        Key := 0;
      end;
    else
      fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;

Initialization
  registerclasses ( [ TOF_AMINTEG ] ) ;
end.
