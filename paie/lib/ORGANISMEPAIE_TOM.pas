{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ORGANISMEPAIE (ORGANISMEPAIE)
Mots clefs ... : TOM;ORGANISMEPAIE
*****************************************************************}
{
PT1    : 14/10/2003 MF V421 Mise en place état d'impression de la fiche
PT2    : 17/10/2003 VG V_42 En l'absence de code de regroupement, décocher et
                            griser "code DADS-U" - FQ N°10892
PT3    : 01/04/2004 VG V_50 Ajout de contrôles liés au champ POG_PREVOYANCE
                            FQ N°10783
PT4    : 16/04/2004 MF V_50 FQ 11045 : Type organisme alpha possible
PT5    : 06/09/2004 PH V_50 FQ 11527 : Mauvais focus
PT6    : 18/11/2004 MF V_60 FQ 11177 : Rupture groupe interne impossible qd code
                                       institution non renseigné.
PT7    : 25/07/2005 MF V_604 FQ 11044 : affichage des codes institution y
                             compris les G
PT8    : 25/07/2005 MF V_604 FQ 12183 : contrôles paramétrage organisme
PT9    : 30/09/2005 MF V_610 On rend invisibles les champs posés pour la
                             DUCS V4.2
PT10   : 20/10/2005 SB V_65 FQ 12273 Refonte Suppression modélisations organismes
PT11   : 09/02/2006 MF V_65 DUCS EDI V4.2
PT12   : 13/02/2006 MF V_65 FQ 11793 (IRC) Ajout coche conditions spéciales de
                            cotitation
PT13   : 27/03/2006 MF V_65 FQ 12542 Qd on ferme (OnClose) la fiche contrôle
                            présence siret organisme + message d'alerte
PT14   : 28/03/2006 MF V_65 FQ 12918 création organisme en utilisant le bouton
                            Annuaire
PT15   : 02/05/2006 MF V_65 FQ 12183 amélioration alimentation identifiants
PT16   : 07/07/2006 MF V_65 FQ 13360 correction erreur SQl qd création à partir
                            annuaire
PT17   : 15/09/2006 VG V_70 Pour les organismes de type BTP, affichage des
                            institutions BTP - FQ N°12982
PT18   : 29/01/2007 MF V_70 modifs DUCS V 4.2
PT19   : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT20   : 13/04/2007 MF V_72 FQ 13743 - modif contrôle type de bordereau
PT21   : 05/06/2007 MF V_72 FQ 14168 - anomalie CWAS
PT22   : 18/10/2006 MF V_80 FQ 14870 ajout coche EDI pour contrôle champs spécifique EDI présents
PT24   : 23/05/2008 FC V_850 FQ 15354 Duplication des organismes collecteurs
}
Unit ORGANISMEPAIE_TOM ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     HDB,
     EdtREtat, // PT1
     Fiche,//PT24
{$ELSE}
     MaineAgl,
     UtileAGL,
     eFiche,//PT24
{$ENDIF}
     UTob,
     sysutils,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     ParamSoc,
     PgOutils2,
// unused     P5Def,
     stdctrls, // PT22
     UTOM,
     Commun,
     Constantes;

Type
  TOM_ORGANISMEPAIE = Class (TOM)
    private
    SiretEmett,qualifiant                : string;
    EDI                                  : TCheckBox;  // PT22
{$IFDEF DUCS41}
    procedure EnterPaieMode (Sender: TObject);
{$ENDIF}
    procedure OnClickAnnuaire(Sender: TObject);
    procedure OnClickImprimer(Sender: TObject);
    procedure OnClickDupliquer(Sender: TObject); //PT24
    procedure ChangeCocheEDI(Sender: TObject);  // PT22
// d PT20
{$IFNDEF DUCS41}
    Function  AffectTypeBordereau( NatureOrg, Periodicite : string; DucsDossier : boolean):string;
    Function  VerifTypeBordereau( NatureOrg, Periodicite : string;  TypePeriod  : string; DucsDossier : boolean):boolean;
{$ENDIF}
// f PT20

    public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation
uses Pgoutils;

{$IFNDEF DUCS41}
// d PT11 DUCS V4.2

{***********A.G.L.***********************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/02/2006
Modifié le ... :   /  /
Description .. : fonction AffectTypeBordereau
Suite ........ : Alimentation du type de bordereau en fonction de la nature
Suite ........ : ducs, de la périodicité et du booléen DUCSDOSSIER
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES, DUCSEDI_V42
*****************************************************************}
// d PT20 Function  AffectTypeBordereau( NatureOrg, Periodicite : string;  TypePeriod  : string; DucsDossier : boolean):string;
Function  TOM_ORGANISMEPAIE.AffectTypeBordereau( NatureOrg, Periodicite : string; DucsDossier : boolean):string;
var TypePeriod : string;
// f PT20
begin
  TypePeriod := '';  // PT20
  result := '';
    if (NatureOrg = '100') then
    // ACOSS
    begin
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (not DucsDossier) then
        TypePeriod := '913';    // BRC d'un établissement
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (DucsDossier) then
        TypePeriod := '914';    // BRC de plusieurs établissements
      if (Periodicite = 'A') and (not DucsDossier) then
        TypePeriod := '915';    // TR d'un établissement
      if (Periodicite = 'A') and (DucsDossier) then
        TypePeriod := '916'     // TR de plusieurs établissements
    end
    else
    if (NatureOrg = '200') then
    // UNEDIC
    begin
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (not DucsDossier) then
        TypePeriod := '920';    // ADV d'un établissement
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (DucsDossier) then
        TypePeriod := '921';    // ADV de plusieurs établissements
      if (Periodicite = 'A') and (not DucsDossier) then
        TypePeriod := '922';     // DRA d'un établissement
      if (Periodicite = 'A') and (DucsDossier) then
        TypePeriod := '923'     // DRA de plusieurs établissements
    end
    else
    if (NatureOrg = '300') then
    // IRC
    begin
      if (Periodicite = 'M') then
        TypePeriod := '931';    // Bordereau mensuel
      if (Periodicite = 'T') then
        TypePeriod := '930';    // Bordereau trimestriel
      if (Periodicite = 'A') then
        TypePeriod := '932';    // Bordereau annuel
    end;
  Result := TypePeriod;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/02/2006
Modifié le ... :   /  /
Description .. : fonction VerifTypeBordereau
Suite ........ : contrôle du type de borderau choisi en fonction de la nature
Suite ........ : ducs.
Mots clefs ... : PAIE, PGDUCS,PGORGANISME,DUCSEDI_V42
*****************************************************************}
Function  TOM_ORGANISMEPAIE.VerifTypeBordereau( NatureOrg, Periodicite : string;  TypePeriod  : string; DucsDossier : boolean):boolean;
begin
     result := true;
     if ((NatureOrg = '100') and ((TypePeriod <> '913') and (TypePeriod <> '914')
          and (TypePeriod <> '915') and (TypePeriod <> '916'))) then
       result := false;
     if ((NatureOrg = '200') and ((TypePeriod <> '920') and (TypePeriod <> '921')
          and (TypePeriod <> '922') and (TypePeriod <> '923') and (TypePeriod <> '924')
          and (TypePeriod <> '926'))) then
       result := false;
     if ((NatureOrg = '300') and ((TypePeriod <> '931') and (TypePeriod <> '930')
          and (TypePeriod <> '932'))) then
       result := false;
     if (((Periodicite = 'M') OR  (Periodicite = 'T')) and
         ((TypePeriod = '915') or (TypePeriod = '916') or (TypePeriod = '922') or
          (TypePeriod = '924') or (TypePeriod = '923') or (TypePeriod = '926') or
          (TypePeriod = '932'))) then
       result := false;
     if ((Periodicite = 'A')  and
         ((TypePeriod <> '915') and (TypePeriod <> '916') and (TypePeriod <> '922') and
          (TypePeriod <> '924') and (TypePeriod <> '923') and (TypePeriod <> '926') and
          (TypePeriod <> '932'))) then
       result := false;
     if ((Periodicite <> 'M') and (TypePeriod = '931')) then
       result := false;
     if ((Periodicite <> 'T') and (TypePeriod = '930')) then
       result := false;
     if ((Periodicite <> 'A') and (TypePeriod = '932')) then
       result := false;
     if ((DucsDossier) and ((TypePeriod = '913') or (TypePeriod = '915') or
         (TypePeriod = '920') or (TypePeriod = '922') or (TypePeriod = '924'))) then
       result := false;
     if ((not DucsDossier) and ((TypePeriod = '914') or (TypePeriod = '916') or
         (TypePeriod = '921') or (TypePeriod = '923') or (TypePeriod = '926'))) then
       result := false;
end;
{$ENDIF}

// f PT11 DUCS V4.2

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : procédure OnNewRecord.
Suite ........ : initialisation des champs de l'enregistrement
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
procedure TOM_ORGANISMEPAIE.OnNewRecord ;
begin
  Inherited ;
  {affectation des valeurs par défaut}
  SetField('POG_LIBEDITBULL','');
  SetField('POG_NUMAFFILIATION','');
  SetField('POG_PERIODICITDUCS','');
  SetField('POG_CAISSDESTIN','-');
  SetField('POG_AUTREPERIODUCS','');
  SetField('POG_INSTITUTION','');
  SetField('POG_NUMINTERNE','');
  SetField('POG_LGOPTIQUE','');
  SetField('POG_LONGTOTALE','7');
  SetField('POG_LONGEDITABLE','0');
  SetField('POG_POSDEBUT','0');
  SetField('POG_BASETYPARR','P');
  SetField('POG_MTTYPARR','P');
  SetField('POG_TELEPHONE','');
  SetField('POG_FAX','');
  SetField('POG_TELEX','');
  SetField('POG_SIRET','');
  SetField('POG_CONTACT','');
  SetField('POG_REGROUPEMENT','');
  SetField('POG_RUPTSIRET','-');
  SetField('POG_RUPTAPE','-');
  SetField('POG_RUPTGROUPE','-');
  SetField('POG_REGROUPDADSU','-');
  SetField('POG_PREVOYANCE','-');  //PT3
  SetField('POG_EMAIL','');
  SetField('POG_RUPTNUMERO','-');
  SetField('POG_ADRESSE1','');
  SetField('POG_ADRESSE2','');
  SetField('POG_ADRESSE3','');
  SetField('POG_CODEPOSTAL','');
  SetField('POG_VILLE','');
  SetField('POG_GENERAL','');
//PT22  SetField('POG_CODAPPLI','');
  SetField('POG_CODAPPLI','X  ');
  SetField('POG_SERVUNIQ','-');
  SetField('POG_SOUSTOTDUCS','-');
  SetField('POG_POSTOTAL','0');
  SetField('POG_LONGTOTAL','0');
  SetField('POG_PAIEGROUPE','-');
  SetField('POG_RIBDUCSEDI','');
  SetField('POG_PAIEMODE','');
  SetField('POG_IDENTOPS','');
  SetField('POG_NOCONTEMET','');
  SetField('POG_DUCSDOSSIER','-');
  SetField('POG_PERIODCALCUL','X');
  SetField('POG_AUTPERCALCUL','X');
  SetField('POG_DUCSREGLEMENT','0');
  SetField('POG_DUCEXIGIBILITE','0');
  SetField('POG_DUCLIMITEDEPOT','0');
  SetField('POG_VCPA','-');
//PT15  SetField('POG_IDENTQUAL','5');
  SetField('POG_IDENTQUAL',' ');
  SetField('POG_IDENTEMET','');
  SetField('POG_IDENTDEST','');
  SetField('POG_ADHERCONTACT','');
  SetField('POG_CENTREPAYEUR','');
  SetField('POG_CPINFO','');
  SetField('POG_CONDSPEC','-');      //PT12
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : Procédure OnDeleteRecord
Suite ........ : Suppression de l'organisme, vérification qu'aucune rubrique
Suite ........ : de cotisation attachée à cet organisme n'existe dans
Suite ........ : l'historique
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
procedure TOM_ORGANISMEPAIE.OnDeleteRecord ;
var
  Q                             : TQuery ;
  St                            : string ;
  Nb                            : integer ;
  CEG, STD, DOS                 : Boolean ;
begin
  inherited;
  {Une rubrique de cotisation associée à cet organisme pour cet établissement
   a-t-elle déjà été utilisée ? }
  st := 'SELECT Count (*) NBRE FROM HISTOBULLETIN '+
        'WHERE PHB_ETABLISSEMENT = "' + GetField('POG_ETABLISSEMENT') + '" AND '+
        'PHB_ORGANISME = "'+ GetField('POG_ORGANISME') + '" AND '+
        'PHB_NATURERUB = "COT"' ;
  Q := OpenSql (st, TRUE) ;
  if not Q.EOF then
    nb := Q.FindField ('NBRE').AsInteger
  else
    nb := 0;
  Ferme (Q);

  if nb <> 0 then
  {suppression impossible}
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Certaines cotisations alimentent cet organisme,'+
                  '#13#10Vous ne pouvez le supprimer!';
  end
  else
  {suppression possible - La suppression de la ventilation associée est-elle possible?}
  {le même organisme existe-t'il pour un autre établissement?}
  begin
    st := 'SELECT Count (*) NBRE FROM ORGANISMEPAIE '+
        'WHERE POG_ETABLISSEMENT <> "' + GetField('POG_ETABLISSEMENT') + '" AND '+
        'POG_ORGANISME = "'+ GetField('POG_ORGANISME') + '"';
    Q := OpenSql (st, TRUE);
    if not Q.EOF then
      nb := Q.FindField ('NBRE').AsInteger
    else
      nb := 0;
    Ferme (Q);

    { La suppression de la ventilation est possible}
    if (nb = 0) then
      { DEB PT10 }
      if PgiAsk('Voulez-vous supprimer les modélisations comptables liées à ce type d''organisme?',Ecran.caption) = MrYes then
      Begin
      AccesPredefini('TOUS', CEG, STD, DOS);
      if CEG = True then
        ExecuteSQL('DELETE FROM VENTIORGPAIE WHERE ##PVO_PREDEFINI## PVO_TYPORGANISME="'+GetField('POG_ORGANISME')+'" AND PVO_PREDEFINI="CEG"');
      if STD = True then
        ExecuteSQL('DELETE FROM VENTIORGPAIE WHERE ##PVO_PREDEFINI## PVO_TYPORGANISME="'+GetField('POG_ORGANISME')+'" AND PVO_PREDEFINI="STD"');
      if STD = True then
        ExecuteSQL('DELETE FROM VENTIORGPAIE WHERE ##PVO_PREDEFINI## PVO_TYPORGANISME="'+GetField('POG_ORGANISME')+'" AND PVO_PREDEFINI="DOS"');
      End;
      { FIN PT10 }
  end;
end ;

procedure TOM_ORGANISMEPAIE.OnUpdateRecord ;
var
{$IFDEF EAGLCLIENT}
  ControlOrg : THValComboBox;
{$ELSE}
  ControlOrg : THDBValComboBox;
{$ENDIF}
  st,ValLibelle                         : String;

{$IFNDEF DUCS41}
  VNatureOrg,VPeriodicitDucs, VTypePeriod, VAutrePerioDucs,VTypeAutPeriod : string; //PT11 DUCS EDI V4.2
  VDucsDossier                                                            : boolean; //PT11 DUCS EDI V4.2
{$ENDIF} 

begin
  inherited;
st:='';
{$IFDEF EAGLCLIENT}
ControlOrg:=THValComboBox(GetControl('POG_ORGANISME'));
{$ELSE}
  ControlOrg:=THDBValComboBox(GetControl('POG_ORGANISME'));
{$ENDIF}
  if ControlOrg <> NIL then
  begin
    ValLibelle:= RechDom('PGTYPEORGANISME',ControlOrg.Value,FALSE) ;
    if (ValLibelle='') or (ValLibelle='Error') then
      LastError:=1;
    if LastError <> 0 then
    begin
      SetFocusControl('POG_ORGANISME');
      LastErrorMsg:='Le code organisme est obligatoirement renseigné';
    end;
  end;
  st := GetField ('POG_ETABLISSEMENT');
  if st = '' then
  begin
    SetFocusControl('POG_ETABLISSEMENT');
    LastErrorMsg:='Le code POG_ETABLISSEMENT est obligatoirement renseigné';
  end;

  if (GetField ('POG_PAIEGROUPE')='X') and (GetField('POG_DUCSDOSSIER')<>'X') then
  begin
    SetFocusControl('POG_PAIEGROUPE');
    LastErrorMsg:='Ce n''est pas une ducs dossier, paiement groupé impossible';
    LastError:=1
  end;
// d PT6
  if (GetField ('POG_RUPTGROUPE')='X') and (GetField('POG_INSTITUTION')='') then
  begin
    SetFocusControl('POG_RUPTGROUPE');
    LastErrorMsg:='La rupture sur groupe interne n''est pas possible.';
    LastError:=1
  end;
// f PT6

{$IFNDEF DUCS41}
// d PT11 DUCS EDI V4.2
// contrôle du champ Titulaire du compte dans le cas du télérèglement pour toutes les natures d'organisme
// et pour les IRC du virement, du prélèvement et du télérèglement
// (champ obligatoire)
//PT18  if (GetField('POG_PAIEMODE') =  'Z10') and (getField('POG_TITULAIRECPT') = '') then
  VNatureOrg := GetField('POG_NATUREORG');
  if (((GetField('POG_PAIEMODE') = 'Z10') or
       ((VNatureOrg = '300') and
        ((GetField('POG_PAIEMODE') = '30') or
         (GetField('POG_PAIEMODE') = '31')))) and
       (getField('POG_TITULAIRECPT') = '')) then
  begin
    SetFocusControl('POG_TITULAIRECPT');
    LastErrorMsg:='Le titulaire du compte doit être renseigné.';
    LastError:=1
  end;
{$ENDIF}

  if (GetField('POG_PAIEMODE') =  'Z10') and (getField('POG_IDENTOPS') = '') then
  begin
    SetFocusControl('POG_IDENTOPS');
    LastErrorMsg:='L''identification de l''OPS doit être renseignée.';
    LastError:=1
  end;
  if (GetField('POG_PAIEMODE') =  '31') and (getField('POG_NATUREORG') <> '300') then
  begin
    SetFocusControl('POG_PAIEMODE');
    LastErrorMsg:='Ce mode de règlement n''est pas utilisable pour cette nature de DUCS.';
    LastError:=1
  end;

{$IFNDEF DUCS41}
  // banque obligatoire pour virement, prélèvement et télérèglement
    if ((GetField('POG_PAIEMODE') =  '30') or (GetField('POG_PAIEMODE') =  '31')
         or (GetField('POG_PAIEMODE') =  'Z10')) and
         (getField('POG_RIBDUCSEDI') = '') then
  begin
    SetFocusControl('POG_RIBDUCSEDI');
    LastErrorMsg:='La banque doit être renseignée.';
    LastError:=1
  end;
  if ((GetField('POG_SIRET') =  '') and (GetField('POG_IDENTDEST') ='')) then // PT13
{$ELSE}
  if (GetField('POG_SIRET') =  '') then
{$ENDIF}
  begin
    if (EDI <> NIL) and (EDI.Checked) then //PT22
      PGIBox('Le siret de l''organisme, nécessaire pour l''EDI, n''est pas renseigné.', 'Attention');     //PT22
  end;

  if (GetField('POG_IDENTQUAL') = '5') then
  begin
    if ((GetField('POG_IDENTEMET') = '') or (GetField('POG_IDENTEMET') = ' '))then
      SetField ('POG_IDENTEMET',SiretEmett);
    if ((GetField('POG_IDENTDEST') = '') or (GetField('POG_IDENTDEST') = ' '))then
      SetField ('POG_IDENTDEST',GetField('POG_SIRET'));
  end;
  if (GetField('POG_NATUREORG') = '200') and (GetField('POG_IDENTQUAL') = 'ZZZ') then
  begin
    if (GetField('POG_IDENTEMET') = '') then
      SetField ('POG_IDENTEMET','ES'+SiretEmett);
    if (GetField('POG_IDENTDEST') = '') then
      SetField ('POG_IDENTDEST','H02S000DUCS');
  end;
// f PT8 FQ 12183


{$IFNDEF DUCS41}
// Contrôle des types de bordereau choisis
//PT18  VNatureOrg := GetField('POG_NATUREORG');
  VPeriodicitDucs := GetField('POG_PERIODICITDUCS');
  VTypePeriod := GetField('POG_TYPEPERIOD');
  VAutrePerioDucs := GetField('POG_AUTREPERIODUCS');
  VTypeAutPeriod := GetField('POG_TYPEAUTPERIOD');
  if (GetField('POG_DUCSDOSSIER')= 'X') then
    VDucsDossier := true
    else
    VDucsDossier := false;

// d PT20
  if (VPeriodicitDucs <> '') then
  begin
    if (EDI <> NIL) and (EDI.Checked) then //PT22
      if (not VerifTypeBordereau(VNatureOrg,VPeriodicitDucs,VTypePeriod,VDucsDossier)) then
        PGIBox('Le type de bordereau choisi est peut-être erroné.', 'Attention');  //PT22
  end;
  if (VAutrePerioDucs <> '') then
  begin
    if (EDI <> NIL) and (EDI.Checked) then //PT22
      if (not VerifTypeBordereau(VNatureOrg,VAutrePerioDucs,VTypeAutPeriod,VDucsDossier)) then
        PGIBox('Le type de bordereau choisi est peut-être erroné.', 'Attention');  //PT22
  end;
  if (VPeriodicitDucs = '') and (VAutrePerioDucs = '') then
    PGIBox('Aucune périodicité choisie pour le bordereau. ', 'Attention'); //PT22
// f PT20

// f PT11 DUCS EDI V4.2
{$ENDIF}
// d PT22
  if (EDI <> NIL) then
  begin
    if (EDI.Checked = True) then
      SetField ('POG_CODAPPLI','X  ')
    else
      SetField ('POG_CODAPPLI','-  ');
  end;
// f PT22
end ;

procedure TOM_ORGANISMEPAIE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : Procédure OnLoad
Suite ........ : affectation de valeur par défaut
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
procedure TOM_ORGANISMEPAIE.OnLoadRecord ;
var
  LigneOptique                          : string;
  i                                     : integer;
// d  PT18
{$IFNDEF DUCS41}
  VNatureOrg                            : string;
{$ENDIF}
// f PT18	
begin
  Inherited ;
  qualifiant :=  GetField('POG_IDENTQUAL');

// d PT18
{$IFNDEF DUCS41}
  VNatureOrg := GetField('POG_NATUREORG');
{$ENDIF}
// f PT18

  if (GetField('POG_REGROUPEMENT') = NULL) then
    Setfield ('POG_REGROUPEMENT','');
// d PT9
{$IFDEF DUCS41}
    SetControlVisible('TPOG_TYPEPERIOD',False);
    SetControlVisible('POG_TYPEPERIOD',False);
    SetControlVisible('POG_TYPEAUTPERIOD',False);
    SetControlVisible('TPOG_TYPEAUTPERIOD',False);
    SetControlVisible('POG_TITULAIRECPT',False);
    SetControlVisible('TPOG_TITULAIRECPT',False);
{$ENDIF}
// f PT9
  {si mode de paiement = Z10 alors il faut renseigner l'identification OPS }
// d PT11 DUCS EDI V4.2
  if (GetField('POG_PAIEMODE') <> 'Z10') then
  begin
    SetControlEnabled('POG_IDENTOPS',false) ;
{$IFNDEF DUCS41}
//PT18    SetControlEnabled('POG_TITULAIRECPT',false);
{$ENDIF}
  end
  else
  begin
    SetControlEnabled('POG_IDENTOPS',true);
{$IFNDEF DUCS41}
//PT18    SetControlEnabled('POG_TITULAIRECPT',true);
{$ENDIF}
  end;
// f PT11 DUCS EDI V4.2

// d PT18
{$IFNDEF DUCS41}
//    SetControlEnabled('POG_TITULAIRECPT',false);
// Le titulaire du compte est à renseigner si télérèglement (Z10) ou pour les
// IRC si virement (30) ou prélèvement  (31)
  if (((GetField('POG_PAIEMODE') <> 'Z10') and (VNatureOrg <> '300')) or
      ((VNatureOrg = '300') and
       ((GetField('POG_PAIEMODE') <> 'Z10') and
        (GetField('POG_PAIEMODE') <> '30') and
        (GetField('POG_PAIEMODE') <> '31')))) then
       SetControlEnabled('POG_TITULAIRECPT',false)
  else
       SetControlEnabled('POG_TITULAIRECPT',true);
{$ENDIF}
// f PT18

  {si IRC alors il faut renseigner le n° de contrat de l'émetteur chez le destinataire}
  if  (GetField('POG_NATUREORG') <> '300') then
  begin
    SetControlEnabled('POG_NOCONTEMET',False);
    SetControlEnabled('GRPCENTREPAYEUR',False);
    SetControlVisible('GRPCENTREPAYEUR',False);
  end
  else
  begin
    SetControlEnabled('POG_NOCONTEMET',True);
    SetControlEnabled('GRPCENTREPAYEUR',True);
    SetControlVisible('GRPCENTREPAYEUR',True);
// d PT12
    SetControlEnabled('POG_CONDSPEC',True);
    SetControlVisible('POG_CONDSPEC',True);
// f PT12
  end;
  {alimentation 30 1er caractères de la ligne optique}
  if (GetField('POG_NUMINTERNE') <> '') and
     (GetField('POG_LGOPTIQUE') = '') then
  begin
    if  (GetField('POG_NATUREORG') <> '200') then
    begin
      LigneOptique := Copy(GetField('POG_NUMINTERNE'),1, length(GetField('POG_NUMINTERNE')));
      if (length(GetField('POG_NUMINTERNE')) < 30) then
      begin
        for i := length(GetField('POG_NUMINTERNE'))+1 to 30 do
        begin
          LigneOptique := LigneOptique + '0';
        end;
      end;
    end
    else
    {ASSEDIC}
    begin
      LigneOptique := 'S2'+Copy(GetField('POG_NUMINTERNE'),1, length(GetField('POG_NUMINTERNE')));
      if ((length(GetField('POG_NUMINTERNE'))+2) < 30) then
      begin
        for i := length(GetField('POG_NUMINTERNE'))+3 to 30 do
        begin
          LigneOptique := LigneOptique + '0';
        end;
      end;
    end;
    SetControlText('POG_LGOPTIQUE',LigneOptique);
  end;
// d PT22
  if (EDI <> NIL) then
  begin
    if ((Copy(GetField('POG_CODAPPLI'),1,1) <> 'X') and
        (Copy(GetField('POG_CODAPPLI'),1,1) <> '-')) or
       (Copy(GetField('POG_CODAPPLI'),1,1) = 'X') then
       EDI.Checked:=True
    else
      if (Copy(GetField('POG_CODAPPLI'),1,1) = '-') then
        EDI.Checked:=False;
    end;
// f PT22
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : Procédure OnChangeField
Suite ........ : Certains champs sont alimenté en fct des valeurs d'autres
Suite ........ : champs.
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
procedure TOM_ORGANISMEPAIE.OnChangeField ( F: TField ) ;
var
{$IFDEF EAGLCLIENT}
NatureOrg                       : THValComboBox;
Champlib                        : THEdit;
{$IFNDEF DUCS41}
ControlBQ                       : THValComboBox;       //PT11 ducs edi V4.2
{$ENDIF}
{$ELSE}
NatureOrg                       : THDBValComboBox;
Champlib                        : THDBEdit;
{$IFNDEF DUCS41}
ControlBQ                       : THDBValComboBox;       //PT11 ducs edi V4.2
{$ENDIF}
{$ENDIF}
TypeOrganisme                   : string;
LigneOptique, Etab,StRechDom    : string;
i                               : integer;
Q                               : Tquery;
VNatureOrg                      : string; // PT11 PT12
// d PT11 DUCS EDI V4.2
{$IFNDEF DUCS41}
VPeriodicitDucs, VTypePeriod, VAutrePerioDucs,VTypeAutPeriod : string;
VDucsDossier                    : boolean;
{$ENDIF}
// f PT11 DUCS EDI V4.2
begin
inherited;

{$IFDEF EAGLCLIENT}
NatureOrg:= THValComboBox (getcontrol ('POG_NATUREORG'));
ChampLib:= THEdit (GetControl ('POG_INSTITUTION'));
{$IFNDEF DUCS41}
ControlBQ:= THValComboBox (getcontrol ('POG_RIBDUCSEDI'));//PT11 DUCS EDI V4.2
{$ENDIF}
{$ELSE}
NatureOrg:= THDBValComboBox (getcontrol ('POG_NATUREORG'));
ChampLib:= THDBEdit (GetControl ('POG_INSTITUTION'));
{$IFNDEF DUCS41}
ControlBQ:= THDBValComboBox (getcontrol ('POG_RIBDUCSEDI')); //PT11 DUCS EDI V4.2
{$ENDIF}
{$ENDIF}

// d PT18
{$IFNDEF DUCS41}
// d PT21 Pour parer anomalie AGL. Voir fq CBP 11138
{$IFNDEF EAGLCLIENT}
VNatureOrg:= GetField ('POG_NATUREORG');
{$ENDIF}
// f PT21
{$ENDIF}
// f PT18

if (F.FieldName = 'POG_NATUREORG') or
   (F.FieldName = 'POG_ORGANISME') then
{Modification Type organisme ou Nature Ducs}
   begin
{par défaut N° Contrat émetteur et Institution inaccessible}
   SetControlEnabled ('POG_NOCONTEMET', False);
   SetControlEnabled ('POG_INSTITUTION', False);
   if ChampLib <> nil then
{on affine la recherche sur la tablette des institutions}
      begin
      TypeOrganisme:= GetField ('POG_ORGANISME');
      if (TypeOrganisme = '003') then            {AGIRC}
         ChampLib.Plus:= ' AND (PIP_INSTITUTION LIKE "C%"'+
                         ' OR PIP_INSTITUTION LIKE "G%")'+
                         ' AND PIP_INSTITUTION not LIKE "Z%"';
      if (TypeOrganisme = '004') then           {ARRCO}
         ChampLib.Plus:= ' AND (PIP_INSTITUTION LIKE "A%"'+
                         ' OR PIP_INSTITUTION LIKE "G%")'+
                         ' AND PIP_INSTITUTION not LIKE "Z%"';
      if (TypeOrganisme <> '003') and (TypeOrganisme <> '004') and
         ((TypeOrganisme < '991') or (TypeOrganisme > '999')) then
         {Autre que Agirc, Arrco, groupe}
         ChampLib.Plus:= ' AND PIP_INSTITUTION not LIKE "G%"'+
                         ' AND PIP_INSTITUTION not LIKE "Z%"';
      if (TypeOrganisme <> '') and (IsNumeric (TypeOrganisme)) and
         (StrToInt (TypeOrganisme) >= 991) and
         (StrToInt (TypeOrganisme) <= 999) then                {Groupe}
         ChampLib.Plus:= ' AND PIP_INSTITUTION  LIKE "G%"';
      end;

   if (NatureOrg <> NIL) and (GetField ('POG_LONGEDITABLE') = '0') and
      (GetField ('POG_POSDEBUT') = '0') and
      (GetField ('POG_LONGTOTAL') = '0') and
      (GetField ('POG_POSTOTAL') = '0') then
      begin
      if (GetField ('POG_IDENTQUAL') <> '5') then
         SetField ('POG_IDENTQUAL', '5');
      if (GetField ('POG_IDENTEMET') <> SiretEmett) then
         SetField ('POG_IDENTEMET', SiretEmett);
      if (GetField ('POG_IDENTDEST') <> GetField ('POG_SIRET')) then
         SetField ('POG_IDENTDEST', GetField ('POG_SIRET'));
      qualifiant:= '5';
      if (NatureOrg.Value = '100') then      // URSSAF
         begin
         if (GetField ('POG_SOUSTOTDUCS') <> '-') then
            SetField ('POG_SOUSTOTDUCS', '-');
         if (GetField ('POG_LONGEDITABLE') <> '4') then
            SetField ('POG_LONGEDITABLE', '4');
         if (GetField ('POG_POSDEBUT') <> '4') then
            SetField ('POG_POSDEBUT', '4');
         end;
      if (NatureOrg.Value = '200') then    // ASSEDIC
         begin
         if (GetField ('POG_IDENTQUAL') <> 'ZZZ') then
//d PT15
            begin
            SetField ('POG_IDENTQUAL', 'ZZZ');
            SetField ('POG_IDENTEMET', '');
            end;
//f PT15
//d PT11
{$IFDEF DUCS41}
         if (copy (GetField ('POG_IDENTEMET'), 1, 2) <> 'ES') then
            SetField ('POG_IDENTEMET', 'ES');
{$ELSE}
         if (GetField ('POG_IDENTEMET') = '') then
            SetField ('POG_IDENTEMET', 'ES'+SiretEmett);
{$ENDIF}

{$IFDEF DUCS41}
         if (GetField ('POG_IDENTDEST') <> 'H02S000DUCS') then
{$ELSE}
         if (GetField ('POG_IDENTDEST') = '') then
{$ENDIF}
            SetField ('POG_IDENTDEST', 'H02S000DUCS');

         qualifiant:= 'ZZZ';
// f PT11
         if (GetField ('POG_SOUSTOTDUCS') <> 'X') then
            SetField ('POG_SOUSTOTDUCS', 'X');
         if (GetField ('POG_LONGTOTAL') <> '1') then
            SetField ('POG_LONGTOTAL', '1');
         if (GetField ('POG_POSTOTAL') <> '3') then
            SetField ('POG_POSTOTAL', '3');
         if (GetField ('POG_LONGEDITABLE') <> '3') then
            SetField ('POG_LONGEDITABLE', '3');
         if (GetField ('POG_POSDEBUT') <> '5') then
            SetField ('POG_POSDEBUT', '5');
         end;
      if (NatureOrg.Value = '300') then    // IRC
         begin
         if (GetField ('POG_SOUSTOTDUCS') <> 'X') then
            SetField ('POG_SOUSTOTDUCS', 'X');
         if (GetField ('POG_LONGEDITABLE') <> '5') then
            SetField ('POG_LONGEDITABLE', '5');
         if (GetField ('POG_POSDEBUT') <> '3') then
            SetField ('POG_POSDEBUT', '3');
         end;
      if (NatureOrg.Value = '600') then    // MSA
         begin
         if (GetField ('POG_SOUSTOTDUCS') <> 'X') then
            SetField ('POG_SOUSTOTDUCS', 'X');
         if (GetField ('POG_LONGTOTAL') <> '2') then
            SetField ('POG_LONGTOTAL', '2');
         if (GetField ('POG_POSTOTAL') <> '3') then
            SetField ('POG_POSTOTAL', '3');
         if (GetField ('POG_LONGEDITABLE') <> '5') then
            SetField ('POG_LONGEDITABLE', '5');
         if (GetField ('POG_POSDEBUT') <> '3') then
            SetField ('POG_POSDEBUT', '3');
         end;
      if (NatureOrg.Value = '700') then    // BTP
         begin
         if (GetField ('POG_SOUSTOTDUCS') <> 'X') then
            SetField ('POG_SOUSTOTDUCS', 'X');
         if (GetField ('POG_LONGTOTAL') <> '2')then
            SetField ('POG_LONGTOTAL', '2');
         if (GetField ('POG_POSTOTAL') <> '3') then
            SetField ('POG_POSTOTAL', '3');
         if (GetField ('POG_LONGEDITABLE') <> '5') then
            SetField ('POG_LONGEDITABLE', '5');
         if (GetField ('POG_POSDEBUT') <> '3') then
            SetField ('POG_POSDEBUT', '3');
         end;
      end;
   if (NatureOrg.Value = '300') then  // IRC
      begin
      SetControlEnabled ('POG_NOCONTEMET', True);
      SetControlVisible ('POG_NOCONTEMET', True);
      SetControlVisible ('TPOG_NOCONTEMET', True);
      SetControlEnabled ('GRPCENTREPAYEUR', True);
      SetControlVisible ('GRPCENTREPAYEUR', True);
      SetControlEnabled ('POG_INSTITUTION', True);
      end
   else
//PT17
   if (NatureOrg.Value = '700') then  // BTP
      begin
      SetControlEnabled ('POG_INSTITUTION', True);
      ChampLib.Plus:= ' AND (PIP_INSTITUTION LIKE "B%")';
      end
   else
//FIN PT17
      begin
      SetControlEnabled ('POG_NOCONTEMET', False);
      SetControlVisible ('POG_NOCONTEMET', False);
      SetControlVisible ('TPOG_NOCONTEMET', False);
      SetControlEnabled ('GRPCENTREPAYEUR', False);
      SetControlVisible ('GRPCENTREPAYEUR', False);
      if (GetField ('POG_CENTREPAYEUR') <> '') then
         SetField ('POG_CENTREPAYEUR', '');
      if (GetField ('POG_CPINFO') <> '') then
         SetField ('POG_CPINFO', '');
      if (GetField ('POG_NOCONTEMET') <> '') then
         SetField ('POG_NOCONTEMET', '');
      SetControlEnabled ('POG_INSTITUTION', False);
      if (GetField ('POG_INSTITUTION') <> '') then
         SetField ('POG_INSTITUTION', '');
      end;
// d PT11 DUCS EDI V4.2
// affectation par défaut du type de bordereau
   if (F.FieldName = 'POG_NATUREORG') then
      begin
      VNatureOrg:= GetField ('POG_NATUREORG');
{$IFNDEF DUCS41}
      VPeriodicitDucs:= GetField ('POG_PERIODICITDUCS');
      VTypePeriod:= GetField ('POG_TYPEPERIOD');
      if (GetField ('POG_DUCSDOSSIER') = 'X') then
         VDucsDossier:= true
      else
         VDucsDossier:= false;

      VAutrePerioDucs:= GetField ('POG_AUTREPERIODUCS');
      VTypeAutPeriod:= GetField ('POG_TYPEAUTPERIOD');

{ PT20
      VTypeAutPeriod:= AffectTypeBordereau (VNatureOrg, VAutrePerioDucs,
                                            VTypeAutPeriod, VDucsDossier);}
      VTypeAutPeriod:= AffectTypeBordereau (VNatureOrg, VAutrePerioDucs,
                                            VDucsDossier);
      SetField ('POG_TYPEAUTPERIOD', VTypeAutPeriod);

{ PT20
     VTypePeriod:= AffectTypeBordereau (VNatureOrg, VPeriodicitDucs,
                                         VTypePeriod, VDucsDossier);}
      VTypePeriod:= AffectTypeBordereau (VNatureOrg, VPeriodicitDucs,
                                         VDucsDossier);
      SetField ('POG_TYPEPERIOD', VTypePeriod);
{$ENDIF}
// d PT12
      if (VNatureOrg = '300') then             // IRC
         begin
         SetControlEnabled ('POG_CONDSPEC', true);
         SetControlVisible ('POG_CONDSPEC', true);
         end
      else
         begin
         SetControlEnabled ('POG_CONDSPEC', false);
         SetControlVisible ('POG_CONDSPEC', false);
         SetField ('POG_CONDSPEC', '-');
         end;
// f PT12
      end;
// f PT11 DUCS EDI V4.2
   exit;
   end;

if (F.FieldName ='POG_SIRET') then {Contrôle du Siret}
   begin
   if ((DS.State in [dsInsert]) and (GetField ('POG_ORGANISME') <> '') and
      (GetField ('POG_NATUREORG') <> '')) or (((DS.State in [dsBrowse]) or
      (DS.State in [dsEdit])) and (GetField ('POG_SIRET') <> '')) then
      begin
      if (Length (GetField ('POG_SIRET')) <> 14) and
         (Length (GetField ('POG_SIRET')) <> 9) then
         begin
         PGIBox ('! Attention, N° Siret incomplet', 'N° Siret');
         SetFocusControl ('POG_SIRET');
         end
      else
      if (ControlSiret (GetField ('POG_SIRET')) <> True) then
         begin
         PGIBox ('! Attention, N° Siret erroné','N° Siret');
         SetFocusControl ('POG_SIRET');
         end
      else
         begin           {alimentation identifiant destinataire}
{ Siret}
         if (GetField ('POG_IDENTDEST') = '') and
            (GetField ('POG_IDENTQUAL') = '5') then
            if (GetField ('POG_IDENTDEST') <> GetField ('POG_SIRET'))then
               SetField ('POG_IDENTDEST', GetField ('POG_SIRET'));
{Siren}
         if (GetField ('POG_IDENTDEST') = '') and
            (GetField ('POG_IDENTQUAL') = '22') then
            if (GetField ('POG_IDENTDEST') <> Copy (GetField ('POG_SIRET'), 1, 9)) then
               SetField ('POG_IDENTDEST', Copy(GetField ('POG_SIRET'), 1, 9));
         end;
      end;
   exit;
   end;

if F.FieldName = 'POG_IDENTQUAL' then     {traitement qualifiant d'identifiant}
   begin
   if (GetField ('POG_IDENTDEST') = '') then
      begin     {qualifiant destinataire non encore renseigné}
{Siret}
      if (GetField ('POG_IDENTQUAL') = '5') then
         if (GetField ('POG_IDENTDEST') <> GetField ('POG_SIRET')) then
            SetField ('POG_IDENTDEST', GetField ('POG_SIRET'));
{Siren}
      if (GetField ('POG_IDENTQUAL') = '22') then
         if (GetField ('POG_IDENTDEST') <> Copy (GetField ('POG_SIRET'), 1, 9)) then
            SetField ('POG_IDENTDEST', Copy (GetField ('POG_SIRET'), 1, 9));
      if (GetField ('POG_IDENTQUAL') = 'ZZZ') and
         (GetField ('POG_NATUREORG') = '200') then
      {Assedic (par défaut H02S000DUCS si réel, H03S000DUCSsi test)}
{$IFDEF DUCS41}
         if (GetField ('POG_IDENTDEST') <> 'H02S000DUCS') then
{$ELSE}
         if (GetField ('POG_IDENTDEST') = '') then
{$ENDIF}
            SetField ('POG_IDENTDEST', 'H02S000DUCS');
      end
   else
      begin {qualifiant destinataire  déjà renseigné}
{Siren}
      if (GetField ('POG_IDENTQUAL') = '22') then
         begin
         if (qualifiant = '5')  then
{(ancien qualifiant = siret)}
            if (GetField ('POG_IDENTDEST') <> Copy (GetField ('POG_IDENTDEST'), 1, 9)) then
               SetField ('POG_IDENTDEST', Copy (GetField ('POG_IDENTDEST'), 1, 9));

         if (qualifiant = 'ZZZ') then
{(ancien qualifiant = selon accord)}
            if (GetField ('POG_IDENTDEST') <> Copy (GetField ('POG_SIRET'), 1, 9)) then
               SetField ('POG_IDENTDEST', Copy (GetField ('POG_SIRET'), 1, 9));
         end;

{Siret}
      if (GetField ('POG_IDENTQUAL') = '5') then
         if (GetField ('POG_IDENTDEST') <> GetField ('POG_SIRET')) then
            SetField ('POG_IDENTDEST', GetField ('POG_SIRET'));

{selon accord}
      if (GetField ('POG_IDENTQUAL') = 'ZZZ') then
{Assedic (par défaut H02S000DUCS si réel, H03S000DUCSsi test)}
         if (GetField ('POG_NATUREORG') = '200') then
// d PT11 DUCS EDI V4.2
{$IFDEF DUCS41}
            if (GetField ('POG_IDENTDEST') <> 'H02S000DUCS') then
               SetField ('POG_IDENTDEST', 'H02S000DUCS')
            else
            if (GetField ('POG_IDENTDEST') <> '') then
               SetField ('POG_IDENTDEST', '');
{$ELSE}
            if (GetField ('POG_IDENTDEST') = '') then
               SetField ('POG_IDENTDEST', 'H02S000DUCS');
{$ENDIF}
// f¨PT11 DUCS EDI V4.2
      end;

   if (GetField ('POG_IDENTEMET') = '') then
      begin     {qualifiant émetteur non encore renseigné}
{ Siret}
      if (GetField ('POG_IDENTQUAL') = '5') then
         if (GetField ('POG_IDENTEMET') <> SiretEmett) then
            SetField ('POG_IDENTEMET', SiretEmett);
{Siren}
      if (GetField ('POG_IDENTQUAL') = '22') then
         if (GetField ('POG_IDENTEMET') <> Copy (SiretEmett, 1, 9)) then
            SetField ('POG_IDENTEMET', Copy (SiretEmett, 1, 9));

{selon accord}
      if (GetField ('POG_IDENTQUAL') = 'ZZZ') then
         if (GetField ('POG_NATUREORG') = '200') then    {Assedic (par défaut ES....}
// d PT11 DUCS EDI V4.2
{$IFDEF DUCS41}
            if (GetField ('POG_IDENTEMET') <> 'ES') then
               SetField ('POG_IDENTEMET', 'ES')
{$ELSE}
            if (GetField ('POG_IDENTEMET') = '') then
               SetField ('POG_IDENTEMET', 'ES'+SiretEmett)
{$ENDIF}
// PT11 DUCS EDI V4.2
            else
            if (GetField ('POG_IDENTEMET') <> '') then
               SetField ('POG_IDENTEMET', '');
      end
   else
      begin    {qualifiant émetteur déjà renseigné}
{Siret}
      if (GetField ('POG_IDENTQUAL') = '5') then
         if (GetField ('POG_IDENTEMET') <> SiretEmett) then
            SetField ('POG_IDENTEMET', SiretEmett);

{Siren}
      if (GetField ('POG_IDENTQUAL') = '22') then
         begin
         if (qualifiant = '5') then
            if (GetField ('POG_IDENTEMET') <> Copy (GetField ('POG_IDENTEMET'), 1, 9)) then
               SetField ('POG_IDENTEMET', Copy (GetField ('POG_IDENTEMET'), 1, 9));
         if (qualifiant = 'ZZZ') then
            if (GetField ('POG_IDENTEMET') <> Copy (SiretEmett, 1, 9)) then
               SetField ('POG_IDENTEMET', Copy (SiretEmett, 1, 9));
         end;

      if (GetField ('POG_IDENTQUAL') = 'ZZZ') then
         if (GetField ('POG_NATUREORG')='200') then         {Assedic }
//  d PT11 DUCS EDI V4.2
{$IFDEF DUCS41}
            if (GetField ('POG_IDENTEMET') <> 'ES') then
               SetField ('POG_IDENTEMET', 'ES')
            else
            if (GetField ('POG_IDENTEMET') <> '') then
               SetField ('POG_IDENTEMET', '');
{$ELSE}
            if (GetField ('POG_IDENTEMET') = '') then
               SetField ('POG_IDENTEMET', 'ES'+SiretEmett);
{$ENDIF}

// f PT11 DUCS EDI V4.2
      end;
   exit;
   end;

if F.FieldName ='POG_IDENTEMET' then      {traitement identifiant émetteur}
   begin
   if (GetField ('POG_IDENTEMET') <> '') then
      begin
      if (GetField ('POG_IDENTQUAL') = '5') or
         (GetField ('POG_IDENTQUAL') = '22') then
         if (Length (GetField ('POG_IDENTEMET')) <> 14) and
            (Length (GetField ('POG_IDENTEMET')) <> 9) then
            begin
            PGIBox ('! Attention, Siret ou Siren incomplet', 'Siret ou Siren');
            SetFocusControl ('POG_IDENTEMET');
            end
         else
         if ControlSiret (GetField ('POG_IDENTEMET')) <> True then
            begin
            PGIBox ('! Attention, Siret ou Siren erroné', 'Siret ou Siren');
            SetFocusControl ('POG_IDENTEMET');
            end;
      end;
   exit;
   end;

if F.FieldName ='POG_IDENTDEST' then          {traitement identifiant destinataire}
   begin
   if (GetField ('POG_IDENTDEST') <> '') then
      begin
      if (GetField ('POG_IDENTQUAL') = '5') or
         (GetField ('POG_IDENTQUAL') = '22') then
         if (Length (GetField ('POG_IDENTDEST')) <> 14) and
            (Length (GetField ('POG_IDENTDEST')) <> 9) then
            begin
            PGIBox ('! Attention, Siret ou Siren incomplet', 'Siret ou Siren');
            SetFocusControl ('POG_IDENTDEST');
            end
         else
         if ControlSiret (GetField ('POG_IDENTDEST')) <> True then
            begin
            PGIBox ('! Attention, Siret ou Siren erroné', 'Siret ou Siren');
            SetFocusControl ('POG_IDENTDEST');
            end;
      end;
   exit;
   end;

if F.FieldName ='POG_PAIEMODE' then
   begin
{    if (GetField('POG_RIBDUCSEDI') <> '') then
      SetField('POG_RIBDUCSEDI','');}
   Etab:= GetField ('POG_ETABLISSEMENT');
{si mode de paiement = Z10 alors il faut renseigner l'identification OPS}
   if (GetField ('POG_PAIEMODE') <> 'Z10') and
      (GetField ('POG_PAIEMODE') <> '') then
      begin
      SetControlEnabled ('POG_IDENTOPS', false);
{$IFNDEF DUCS41}
//PT18      SetControlEnabled ('POG_TITULAIRECPT', false);
{$ENDIF}
      if (GetField ('POG_IDENTOPS') <> '') then
         SetField ('POG_IDENTOPS', '');
// d PT11 DUCS EDI V4.2
{$IFNDEF DUCS41}
//PT18      if (GetField ('POG_TITULAIRECPT') <> '') then
//PT18         SetField ('POG_TITULAIRECPT', '');
{$ENDIF}
// f PT11 DUCS EDI V4.2
      end
   else
// d PT11 DUCS EDI V4.2
      begin
      SetControlEnabled ('POG_IDENTOPS', true);
{$IFNDEF DUCS41}
//PT18      SetControlEnabled ('POG_TITULAIRECPT', true);
{$ENDIF}
      end;
// f PT11 DUCS EDI V4.2

// d PT18
{$IFNDEF DUCS41}
   if (((GetField ('POG_PAIEMODE') <> 'Z10') and (VNatureOrg <> '300')) or
       ((VNatureOrg = '300') and
        ((GetField('POG_PAIEMODE') <> 'Z10') and
         (GetField('POG_PAIEMODE') <> '30') and
         (GetField('POG_PAIEMODE') <> '31')))) and
      (GetField ('POG_PAIEMODE') <> '') then
       SetControlEnabled ('POG_TITULAIRECPT', false)
   else
       SetControlEnabled ('POG_TITULAIRECPT', true);
{$ENDIF}
// f PT18

   if (GetField ('POG_PAIEMODE') = '30') or
      (GetField ('POG_PAIEMODE') = '31') or
      (GetField ('POG_PAIEMODE') = 'Z10') then       {Virement, Prélèvement, Télérèglement}
      begin
      SetControlEnabled ('POG_RIBDUCSEDI', True);
      if (GetField ('POG_PAIEMODE') = '30') then
{Virement (banque d l'organisme)}
         begin
// d PT11 DUCS EDI V4.2
{$IFNDEF DUCS41}
         if (ControlBQ <> nil) and (ControlBQ.DataType <> 'PGBQORG') then
            SetField ('POG_RIBDUCSEDI', '');
         if (ControlBQ <> nil) then
            ControlBQ.DataType:= 'PGBQORG';
         exit;
{$ENDIF}

{$IFDEF DUCS41}
{$IFDEF EAGLCLIENT}
// PT11 DUCS EDI V4.2
         THValComboBox (GetControl ('POG_RIBDUCSEDI')).DataType:= 'PGBQORG';
{$ELSE}
// PT11 DUCS EDI V4.2
         THDBValComboBox (GetControl ('POG_RIBDUCSEDI')).DataType:= 'PGBQORG';
{$ENDIF}
{$ENDIF}
// f PT11 DUCS EDI V4.2
         end
      else
{Prélèvement ou télérèglement (banque réservée au paiement des charges sociales)}
         begin
// d PT11 DUCS EDI V4.2
{$IFNDEF DUCS41}
         if (ControlBQ <> nil) and (ControlBQ.DataType <> 'TTBANQUECP') then
            SetField ('POG_RIBDUCSEDI', '');
{$ENDIF}
// f PT11 DUCS EDI V4.2

         if (GetField ('POG_RIBDUCSEDI') = '') then
            begin
            Q:= OpenSQL ('SELECT ETB_RIBDUCSEDI'+
                         ' FROM ETABCOMPL WHERE'+
                         ' ETB_ETABLISSEMENT = "'+Etab+'"', True);
            if Not Q.EOF then
               begin
               if (GetField ('POG_RIBDUCSEDI') <> Q.FindField ('ETB_RIBDUCSEDI').AsString) then
                  SetField ('POG_RIBDUCSEDI', Q.FindField ('ETB_RIBDUCSEDI').AsString);
               end;
            ferme (Q);
            end;
// d PT11 DUCS EDI V4.2
{$IFNDEF DUCS41}
         if (ControlBQ <> nil) then
            ControlBQ.DataType:= 'TTBANQUECP';
{$ENDIF}

{$IFDEF DUCS41}
{$IFDEF EAGLCLIENT}
         THValComboBox (GetControl ('POG_RIBDUCSEDI')).DataType:= 'TTBANQUECP';
{$ELSE}
         THDBValComboBox (GetControl ('POG_RIBDUCSEDI')).DataType:= 'TTBANQUECP';
{$ENDIF}
{$ENDIF}
         SetPlusBanqueCP (GetControl ('POG_RIBDUCSEDI'));           //PT19
         end;
      end
   else
// d PT11 DUCS EDI V4.2
      begin
{$IFNDEF DUCS41}
      SetField ('POG_RIBDUCSEDI', '');
{$ENDIF}
      SetControlEnabled ('POG_RIBDUCSEDI', False);
      end;
// f PT11 DUCS EDI V4.2
   exit;
   end;

if F.FieldName ='POG_NUMINTERNE' then
   begin
  {alimentation 30 1er caractères de la ligne optique}
   if (GetField ('POG_NUMINTERNE') <> '') and
      (GetField ('POG_LGOPTIQUE') = '') then
      begin
      if (GetField ('POG_NATUREORG') <> '200') then
         begin
         LigneOptique:= Copy (GetField ('POG_NUMINTERNE'), 1, length (GetField ('POG_NUMINTERNE')));
         if (length (GetField ('POG_NUMINTERNE')) < 30) then
            begin
            for i := length (GetField ('POG_NUMINTERNE'))+1 to 30 do
                LigneOptique:= LigneOptique+'0';
            end;
         end
      else
      {ASSEDIC}
         begin
         LigneOptique:= 'S2'+Copy (GetField ('POG_NUMINTERNE'), 1, length (GetField ('POG_NUMINTERNE')));
         if ((length (GetField ('POG_NUMINTERNE'))+2) < 30) then
            begin
            for i := length (GetField ('POG_NUMINTERNE'))+3 to 30 do
                LigneOptique:= LigneOptique+'0';
            end;
         end;
      if (GetField ('POG_LGOPTIQUE') <> LigneOptique) then
         SetField ('POG_LGOPTIQUE', LigneOptique);
      end;
      exit;
   end;

if (F.FieldName = 'POG_REGROUPEMENT') then
   begin
   if (GetField ('POG_REGROUPEMENT') <> '') then
      begin
      if (GetField ('POG_PREVOYANCE')= 'X') then
         SetControlEnabled ('POG_REGROUPDADSU', False)
      else
         SetControlEnabled ('POG_REGROUPDADSU', True);
      StRechDom:= RechDom ('PGINSTITUTION', GetField ('POG_REGROUPEMENT'), FALSE);
      if (StRechDom = '') or (StRechDom = 'Error') then
         begin
         PGIBox ('Ce code de regroupement ne fait pas partie des institutions',
                 'Code de regroupement');
         if (GetField ('POG_REGROUPEMENT') <> '') then
            SetField ('POG_REGROUPEMENT','');
         SetControlEnabled ('POG_REGROUPDADSU', False);
         SetFocusControl ('POG_REGROUPEMENT');
         end;
      end
   else
      begin
      if (GetField ('POG_REGROUPDADSU')= 'X') then
         SetField ('POG_REGROUPDADSU', '-');
      SetControlEnabled ('POG_REGROUPDADSU', False);
      end;
   exit;
   end;

if (F.FieldName = 'POG_INSTITUTION') then
   begin
   if (not (isnumeric (GetField ('POG_INSTITUTION'))) or
      (length (GetField ('POG_INSTITUTION'))<>4)) then
      begin
      if (GetField ('POG_PREVOYANCE')= 'X') then
         SetField ('POG_PREVOYANCE', '-');
      SetControlEnabled('POG_PREVOYANCE', False);
      end
   else
      begin
      if ((GetField ('POG_REGROUPEMENT') <> '') and
         (GetField ('POG_REGROUPDADSU') = 'X')) then
         SetControlEnabled ('POG_PREVOYANCE', False)
      else
         begin
         SetControlEnabled ('POG_PREVOYANCE', True);
         SetField ('POG_PREVOYANCE', 'X');
         end;
      end;
   if (GetField ('POG_INSTITUTION')='') then
      begin
      SetControlEnabled ('POG_REGROUPEMENT', False);
      SetField ('POG_REGROUPEMENT', '');
      SetControlEnabled ('POG_RUPTGROUPE', False);
      SetField ('POG_RUPTGROUPE', '-');
      end
   else
      begin
      StRechDom:= RechDom ('PGINSTITUTION', GetField ('POG_INSTITUTION'), FALSE);
      if (StRechDom = '') or (StRechDom = 'Error') then
         begin
         PGIBox ('Ce code institution ne fait pas partie des institutions',
                 'Code institution');
         if (GetField ('POG_INSTITUTION') <> '') then
            SetField ('POG_INSTITUTION', '');
         SetControlEnabled ('POG_REGROUPEMENT', False);
         SetFocusControl ('POG_INSTITUTION');
         end
      else
         SetControlEnabled ('POG_REGROUPEMENT', True);

      SetControlEnabled ('POG_RUPTGROUPE', True);
      end;
   exit;
   end;

if (F.FieldName = 'POG_REGROUPDADSU') then
   begin
   if (GetField ('POG_REGROUPDADSU') = 'X') then
      SetControlEnabled ('POG_PREVOYANCE', False)
   else
      begin
      if (not (isnumeric (GetField ('POG_INSTITUTION'))) or
         (length (GetField ('POG_INSTITUTION'))<>4)) then
         SetControlEnabled ('POG_PREVOYANCE', False)
      else
         SetControlEnabled ('POG_PREVOYANCE', True);
      end;
   exit;
   end;

if (F.FieldName = 'POG_PREVOYANCE') then
   begin
   if (GetField ('POG_PREVOYANCE') = 'X') then
      SetControlEnabled ('POG_REGROUPDADSU', False)
   else
      begin
      if (GetField ('POG_REGROUPEMENT') <> '') then
         SetControlEnabled ('POG_REGROUPDADSU', True)
      else
         SetControlEnabled ('POG_REGROUPDADSU', False);
      if (isnumeric (GetField ('POG_INSTITUTION')) and
         (length (GetField ('POG_INSTITUTION'))=4)) then
         begin
         PGIBox ('Incohérence : Cette institution est une prévoyance',
                 'Prévoyance');
         SetField ('POG_PREVOYANCE', 'X');
         end;
      end;
   exit;
   end;

if F.FieldName ='POG_CPINFO' then
   begin
   if (GetField ('POG_CPINFO') <> '') then
      begin
      SetControlEnabled ('POG_CENTREPAYEUR', True);
      SetControlVisible ('POG_CENTREPAYEUR', True);
      SetControlVisible ('TPOG_CENTREPAYEUR',True);
{ N° Affiliation }
      if (GetField ('POG_CPINFO') = '001') then
         begin
         if (GetField ('POG_CENTREPAYEUR') <> GetField ('POG_NUMAFFILIATION')) then
            SetField ('POG_CENTREPAYEUR', GetField ('POG_NUMAFFILIATION'));
         SetControlEnabled ('POG_CENTREPAYEUR', False);
         end;
{ N° Interne }
      if (GetField ('POG_CPINFO') = '002') then
         begin
         if (GetField ('POG_CENTREPAYEUR') <> GetField ('POG_NUMINTERNE')) then
            SetField ('POG_CENTREPAYEUR', GetField ('POG_NUMINTERNE'));
         SetControlEnabled ('POG_CENTREPAYEUR', False);
         end;
{Groupe interne, Siret ou APE (seront alimentés lors de l'initialisation)}
      if (GetField ('POG_CPINFO') = '005') or
         (GetField ('POG_CPINFO') = '003') or
         (GetField ('POG_CPINFO') = '004') then
         begin
         if (GetField ('POG_CENTREPAYEUR') <> '') then
            SetField ('POG_CENTREPAYEUR', '');
         SetControlVisible ('POG_CENTREPAYEUR', False);
         SetControlVisible ('TPOG_CENTREPAYEUR', False);
         end;
      end
   else
      begin
      SetControlVisible ('POG_CENTREPAYEUR', True);
      SetControlVisible ('TPOG_CENTREPAYEUR', True);
      SetControlEnabled ('POG_CENTREPAYEUR', True);
      end;
   exit;
   end;

// d PT11 DUCS EDI V4.2

{$IFNDEF DUCS41}
if (F.FieldName = 'POG_DUCSDOSSIER') then
   begin
   VNatureOrg:= GetField ('POG_NATUREORG');
   VPeriodicitDucs:= GetField ('POG_PERIODICITDUCS');
   VTypePeriod:= GetField ('POG_TYPEPERIOD');
   if (GetField ('POG_DUCSDOSSIER') = 'X') then
      VDucsDossier:= true
   else
      VDucsDossier:= false;
   VAutrePerioDucs:= GetField ('POG_AUTREPERIODUCS');
   VTypeAutPeriod:= GetField ('POG_TYPEAUTPERIOD');

{ PT20
   VTypeAutPeriod:= AffectTypeBordereau (VNatureOrg, VAutrePerioDucs,
                                         VTypeAutPeriod, VDucsDossier);}
   VTypeAutPeriod:= AffectTypeBordereau (VNatureOrg, VAutrePerioDucs,
                                         VDucsDossier);
   SetField ('POG_TYPEAUTPERIOD', VTypeAutPeriod);

{ PT20
   VTypePeriod:= AffectTypeBordereau (VNatureOrg, VPeriodicitDucs, VTypePeriod,
                                      VDucsDossier);}
   VTypePeriod:= AffectTypeBordereau (VNatureOrg, VPeriodicitDucs,
                                      VDucsDossier);
   SetField ('POG_TYPEPERIOD', VTypePeriod);
   exit;
   end;

if (F.FieldName = 'POG_PERIODICITDUCS') then
   begin
   VNatureOrg:= GetField ('POG_NATUREORG');
   VPeriodicitDucs:= GetField ('POG_PERIODICITDUCS');
   VTypePeriod:= GetField ('POG_TYPEPERIOD');
   if (GetField ('POG_DUCSDOSSIER') = 'X') then
      VDucsDossier:= true
   else
      VDucsDossier:= false;

{ PT20
   VTypePeriod:= AffectTypeBordereau (VNatureOrg, VPeriodicitDucs, VTypePeriod,
                                      VDucsDossier);}
   VTypePeriod:= AffectTypeBordereau (VNatureOrg, VPeriodicitDucs,
                                      VDucsDossier);
   SetField ('POG_TYPEPERIOD', VTypePeriod);
   exit;
   end;

if (F.FieldName = 'POG_AUTREPERIODUCS') then
   begin
   VNatureOrg:= GetField ('POG_NATUREORG');

   if (GetField ('POG_DUCSDOSSIER') = 'X') then
      VDucsDossier:= true
   else
      VDucsDossier:= false;

   VAutrePerioDucs:= GetField ('POG_AUTREPERIODUCS');
   VTypeAutPeriod:= GetField ('POG_TYPEAUTPERIOD');

{ PT20
   VTypeAutPeriod:= AffectTypeBordereau (VNatureOrg, VAutrePerioDucs,
                                         VTypeAutPeriod, VDucsDossier);}
   VTypeAutPeriod:= AffectTypeBordereau (VNatureOrg, VAutrePerioDucs,
                                         VDucsDossier);
   SetField ('POG_TYPEAUTPERIOD', VTypeAutPeriod);
//PT11 exit;
   end;
{$ENDIF}
// f PT11 DUCS EDI V4.2
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : Procédure OnArgument
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
procedure TOM_ORGANISMEPAIE.OnArgument ( S: String ) ;
var
  BT, BTI                       : TToolbarButton97; // PT1
  //Pages                         : TPageControl;
  Q                             : Tquery;

{$IFDEF EAGLCLIENT}
{$IFDEF DUCS41}
  PaieMode                      : THValComboBox;
{$ENDIF}
  ChamplibG                     : THEdit;
{$ELSE}
{$IFDEF DUCS41}
  PaieMode                      : THDBValComboBox;
{$ENDIF}
  ChamplibG                     : THDBEdit;
{$ENDIF}
  

begin
  Inherited ;
{$IFDEF CCS3}
  SetControlProperty('PDUCSEDI','TabVisible',False); // Onglet invisible
{$ENDIF}

{$IFDEF EAGLCLIENT}
{$IFDEF DUCS41}
  PaieMode := THValComboBox(GetControl ('POG_PAIEMODE'));
{$ENDIF}
  ChamplibG := THEdit(GetControl('POG_REGROUPEMENT'));
{$ELSE}
{$IFDEF DUCS41}
  PaieMode := THDBValComboBox (GetControl ('POG_PAIEMODE'));
{$ENDIF}
  ChamplibG := THDBEdit(GetControl('POG_REGROUPEMENT'));
{$ENDIF}
//PT11 DUCS EDI V4.2
{$IFDEF DUCS41}
  if PaieMode<>nil then
    PaieMode.OnEnter:=EnterPaieMode;
{$ENDIF}

// d PT22
  EDI := TCheckBox(GetControl('EDI'));
  if EDI <> NIL then
    EDI.OnExit := ChangeCocheEDI;
// f PT22

  if ChampLibG<>nil then
    ChampLibG.Plus:=' AND PIP_INSTITUTION LIKE "G%"';

  BT := TToolbarButton97 (GetControl ('BRECUPANN'));
  if BT <> NIL then
    BT.OnClick := OnClickAnnuaire;
// d PT1
  BTI := TToolbarButton97 (GetControl ('BImprimer'));
  if BTI <> NIL then
    BTI.OnClick := OnClickImprimer;
    BTI.Visible := True;
// f PT1

  //DEB PT24
  BTI := TToolbarButton97 (GetControl ('BDUPLIQUER'));
  if BTI <> NIL then
    BTI.OnClick := OnClickDupliquer;
  //FIN PT24


//PT14  if (PgRendModeFonc ()<> 'MULTI') and (BT <> NIL) then
//PT14   SetControlVisible ('BRECUPANN', FALSE);

{   Pages := TPageControl(GetControl('PAGES'));
 if Pages<>nil then
      Pages.ActivePageIndex:=0;
}
{ Récup. Siret premier émetteur table EMETTEURSOCIAL}
  if (GetParamSocSecur('SO_PGEMETTEUR','') <> '') then   
  {émetteur sociale = celui des paramètres société}
    Q := OpenSQL ('SELECT PET_SIRET '+
                  'FROM EMETTEURSOCIAL '+
                  'WHERE PET_EMETTSOC = "'+
                  GetParamSocSecur('SO_PGEMETTEUR','')+'"', True)  
  else
  {émetteur social = le 1er de la table EMETTEURSOCIAL}
    Q := OpenSQL ('SELECT PET_SIRET '+
                  'FROM EMETTEURSOCIAL ', True);
  SiretEmett := '';
  if Not Q.EOF then
  begin
    SiretEmett := Q.FindField('PET_SIRET').AsString;
  end;
  ferme (Q);


end ;

procedure TOM_ORGANISMEPAIE.OnClose ;
begin
  Inherited ;
// d PT13
  if ((GetField('POG_SIRET') =  '') and (GetField('POG_IDENTDEST') ='') and
      (GetField('POG_ORGANISME') <> '')) then
  begin
    if (EDI <> NIL) and (EDI.Checked) then // PT22
      PGIBox('Le siret de l''organisme, nécessaire pour l''EDI, n''est pas renseigné.', 'Attention');
  end;
// f PT13
end ;

procedure TOM_ORGANISMEPAIE.OnCancelRecord ;
begin
  Inherited ;
end ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... : 21/07/2003
Description .. : Procédure EnterPaieMode
Suite ........ :
Suite ........ : saisie du mode de paiement
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
// PT11 DUCS EDI V4.2
{$IFDEF DUCS41}
procedure TOM_ORGANISMEPAIE.EnterPaieMode (Sender: TObject);
begin
     SetControlEnabled('POG_IDENTOPS',True);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : procédure OnClickAnnuaire
Suite ........ : Récupération des valeur de l'annuaire
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
procedure TOM_ORGANISMEPAIE.OnClickAnnuaire(Sender: TObject);
var
  Retour,st                             : String;
  Q                                     : TQuery ;
begin
  if GetField ('POG_ORGANISME') = '' then exit;

  Retour := AglLanceFiche ('PAY','PGANNUAIRE', '', '' , 'R');

  if (Retour <> '') and (Retour <> 'VIDE') then
  begin
//PT14    St := 'SELECT * FROM ANNUAIRE WHERE ANN_GUIDPER='+Retour;
{PT16    St := 'SELECT POG_LIBELLE,POG_ADRESSE1,POG_ADRESSE2,POG_ADRESSE3,'+
          'POG_CODEPOSTAL,POG_VILLE,POG_TELEPHONE,POG_FAX,POG_EMAIL,POG_SIRET'+
          ' FROM ANNUAIRE WHERE ANN_GUIDPER="'+Retour+'"';}
          St := 'SELECT ANN_NOM1,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,'+
          'ANN_ALCP,ANN_ALVILLE,ANN_TEL1,ANN_FAX,ANN_EMAIL,ANN_SIREN,ANN_CLESIRET'+
          ' FROM ANNUAIRE WHERE ANN_GUIDPER="'+Retour+'"';
    Q := OpenSql (st, TRUE);
    if not Q.EOF then
    begin
      DS.edit;
//PT14      SetField ('POG_LIBELLE',Q.FindField ('ANN_NOMPER').AsString);
      SetField ('POG_LIBELLE',Q.FindField ('ANN_NOM1').AsString);
      SetField ('POG_ADRESSE1', Q.FindField ('ANN_ALRUE1').AsString);
      SetField ('POG_ADRESSE2', Q.FindField ('ANN_ALRUE2').AsString);
      SetField ('POG_ADRESSE3', Q.FindField ('ANN_ALRUE3').AsString);
      SetField ('POG_CODEPOSTAL',Q.FindField ('ANN_ALCP').AsString);
      SetField ('POG_VILLE', Q.FindField ('ANN_ALVILLE').AsString);
      SetField ('POG_TELEPHONE',Q.FindField ('ANN_TEL1').AsString);
      SetField ('POG_FAX', Q.FindField ('ANN_FAX').AsString);
      SetField ('POG_EMAIL',Q.FindField ('ANN_EMAIL').AsString);
      SetField ('POG_SIRET',Q.FindField ('ANN_SIREN').AsString + Q.FindField ('ANN_CLESIRET').AsString);
    end;
    ferme (Q);
  end;
end;
// d PT1
procedure TOM_ORGANISMEPAIE.OnClickImprimer(Sender: TObject);
begin
//PT19
SetControlText ('TRIBDUCSEDI_BQCP', RechDom ('TTBANQUECP',
                GetControlText ('POG_RIBDUCSEDI'), False,
                'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE));
//FIN PT19

LanceEtat('E','PAY','ORG',True,False,false,NIL,
          'AND POG_ETABLISSEMENT = "'+Getfield('POG_ETABLISSEMENT')+'" '+
          'AND POG_ORGANISME ="'+GetField('POG_ORGANISME')+'"' ,'',False)
end;
// f PT1
// d PT22
procedure TOM_ORGANISMEPAIE.ChangeCocheEDI(Sender: TObject);
begin
  if not (ds.state in [dsinsert, dsedit]) then ds.edit;
end;
// f PT22

//DEB PT24
procedure TOM_ORGANISMEPAIE.OnClickDupliquer(Sender: TObject);
var
  AncValCode, AncValEtab : String;
  Champ : array[1..2] of Hstring;
  Valeur : array[1..2] of variant;
  Ok : Boolean;
begin
  TFFiche(Ecran).BValider.Click;
  AncValCode := GetField('POG_ORGANISME');
  AncValEtab := GetField('POG_ETABLISSEMENT');
  AglLanceFiche('PAY', 'DUPLI_ORG', '', '', AncValCode + ';' + AncValEtab);

  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'POG_ORGANISME';
    Valeur[1] := PGCodeDupliquer;
    Champ[2] := 'POG_ETABLISSEMENT';
    Valeur[2] := PGLibDupliquer;
    Ok := RechEnrAssocier('ORGANISMEPAIE', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
      DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('POG_ORGANISME', PGCodeDupliquer);
      SetField('POG_ETABLISSEMENT', PGLibDupliquer);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('BDelete', True);
      SetControlEnabled('BDUPLIQUER', True);
      SetControlEnabled('POG_ORGANISME', False);
      SetControlEnabled('POG_ETABLISSEMENT', False);
      TFFiche(Ecran).BValider.Click;
    end
    else
      HShowMessage('5;Organisme :;La duplication est impossible, l''organisme existe déjà.;W;O;O;O;;;', '', '');
  end;
end;
//FIN PT24

Initialization
  registerclasses ( [ TOM_ORGANISMEPAIE ] ) ;
end.
