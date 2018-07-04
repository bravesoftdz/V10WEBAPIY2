{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 15/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGASSECIC_SPECTACLE
*****************************************************************
PT1 : 18/12/2001 : JL 571 ajout de contrôle avant impressions
PT2 : 27/03/2002 : JL 571 modification complète de la tof :
                   - gestion du navigateur
                   - gestion de la reprise des périodes de travail depuis les paies effectués et les contrats de travail
                   - reprise des zones de la table ETABCOMPL
PT3 : 03/05/2002 : JL 571  Modification du calcul du salaire
PT4 : 21/05/2003 : JL+PH V_42 Portage CWAS le 01/07/03
PT5 : 03/11/2004 : JL V_60 Conversion de type variant incorrect
PT6 : 02/12/2004 : JL V_60 Boutons inaccessibles si aucune période
PT7 : 14/12/2004 : JL V_60 FQ 11621 Taux non renseigné
PT8 : 20/01/2005 : JL V_60 NIR = numéroSS + gestion incrément sur 6 caractères si alpha
PT9 : 04/02/2005 : JL V_60 FQ FQ11832 Libellé forme juridique au lieu du code
PT10 : 04/03/2005 : JL V_60 FQ 11925 Nouveau parametre société pour n° attestation
PT11 : 11/04/2005 : JL V_60 FQ 12176 Correction pour navigateur
PT12 : 04/08/2005 : JL v_60 FQ 12436 Incrémentation du n° d'attestation lors de l'impression au lieu de l'ouverture
PT13 : 09/08/2005 : JL V_60 FQ 12385 Alimentation nb jours + FQ 12640
PT14 : 24/03/2006 : JL V_65 FQ 12850 Gestion salarié cadre
PT15 : 19/05/2006 : JL V_65 Reprise des dates de contrats au lieu des paies
PT16 : 30/08/2006 : JL V_70 FQ 13472 Ajout impression motif fin contart
PT17 : 09/10/2006 : JL V_70 Gestion des réalisateurs
PT18 ! 10/10/2006 : JL V_70 Libellé emploi repris dans contrat
PT19 : 08/12/2006 : JL V_70 reprise date entrée sortie fiche salarié si pas de contrat
PT20 : 24/05/2007 : FL V_72 FQ 13920 Traitement identique des heures et cachets pour techniciens et ouvriers
PT21 : 31/05/2007 : FL V_72 FQ 14029 Pas de motif de fin de contrat si sontrat en cours
PT22 : 31/07/2007 : JL V_80 N° affiliation congés spectacle non accessible
PT23 : 09/08/2007 : JL V_80 FQ 13986 Correction affichage zones à l'impression
PT24 : 09/08/2007 : JL V_80 FQ 13921 Modif pour prendre en compte nom de jeune fille
PT25 : 10/08/2007 : JL V_80 FQ 14405 Gestion paramsoc rubrique cachets
PT26 : 10/09/2007 : RM V_80 Mise à jour pour prendre en compte le formulaire V4
       ==> le 10/09/2007 dev rapide pour un état en vue de l'homologation
       ==> le 05/10/2007 Suite et (normalement) fin du Dev
PT27 : 04/12/2007 : RM-VG V_80 Messages en vision SAV
PT28 : 25/03/2008 : RM V_80 FQ15316 Gestion du Numéro d'objet AEM
PT29 : 26/03/2008 : RM V_80 FQ14017 Correction sur le code emploi
}
unit UTOFPGASSEDIC_spectacle;

Interface
uses     UTOF,
         Controls,
{$IFNDEF EAGLCLIENT}
         DBCtrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
         EdtREtat,
{$ELSE}
       UtileAGL,MaineAgl,P5DEF,
{$ENDIF}
         Hctrls,ComCtrls,HEnt1,HMsgBox,StdCtrls,Classes,
         sysutils,UTob,HTB97,Vierge,ParamDat,Menus,ParamSoc,PGOutils2,UTobDebug;
Type
  TOF_PGASSEDIC_SPECTACLE = Class (TOF)
    procedure OnClose                           ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    private
    //PT26 BImprimer,SalPrem,SalPrec,SalSuiv,SalDern,PerSuiv,PerPrec:TToolbarButton97;
    BImprimer,PerSuiv,PerPrec:TToolbarButton97;
    LeSalarie,EtabSalarie,LeNumAEM,Action:String;
    DateDebutAttest,DateFinAttest:TDateTime;
    NumPeriode,maxPeriode:Integer;
    TobLesPeriodes : Tob;
    TauxEltNat : Double;
    MenuPeriodes : TPopUpMenu;
    LaCategorie,EmploiSalarie : String;
    procedure impression(Sender:TObject) ;
    procedure CalculCotisation(Sender:TObject) ;
    procedure ChargeZonesEmployeur;
    procedure DateElipsisclick(Sender: TObject);
    procedure MetABlanc(ZoneSal,ZonePeriode : Boolean);
    Procedure RemplirTobLesPeriodes;
    Function CalculModulo97(NumAttest : String) : Integer;
    Procedure GereCheckBox (Sender : TObject);
    Procedure AfficherSalarie;
    Procedure AfficherPeriode(NumFille : Integer);
    procedure AfficheDeclarant(Sender : Tobject);
    Procedure PeriodeSuivante(Sender : TObject);
    Procedure PeriodePrecedente(Sender : TObject);
    Procedure MenuPeriodesClick(Sender : TObject);
    procedure ChangeMoisAnnee (Sender : TObject);
    procedure AccesNumAttest(Sender : TObject);
    procedure ValidationAttest (Sender : Tobject);
    procedure RecupAttestation;
    procedure SuppAttest(Sender : TObject);
    Function CodeAnnee(An : Integer) : String;
    procedure RecupNumAttestInit;
    procedure GererCachets(Sender : TObject);  //PT26
    //PT28 Function CtrlSaisie : Boolean;      //PT26
    Function CtrlSaisie(MajBase : Boolean) : Boolean; //PT28
    //procedure RecupDerNumObjet;                //PT26 PT28
end;

Implementation

procedure TOF_PGASSEDIC_SPECTACLE.OnClose;
begin
Inherited ;
    TobLesPeriodes.Free;
end;

procedure TOF_PGASSEDIC_SPECTACLE.OnArgument (stArgument : String ) ;
var DateEdition,DefautTH,Edit : THEdit;
    Q : TQuery;
    Check : TcheckBox;
    Combo : THvalComboBox;
    BVal,BSupp : TtoolBarButton97;
    MoisAttest,AnneeAttest : String;
    aa,mm : Word;
begin
Inherited ;
LeSalarie :=(ReadTokenPipe(stArgument,';'));
LeNumAEM :=(ReadTokenPipe(stArgument,';'));
If LeNumAEM <> '' then Action := 'MODIFICATION'
else Action := 'CREATION';
if Action = 'MODIFICATION' then
begin
     TFVierge(Ecran).Caption := 'Modification de l''attestation assedic spectacle';
     UpdateCaption(TFVierge(Ecran));
end;

MoisAttest := ReadTokenPipe(stArgument,';');
AnneeAttest := ReadTokenPipe(stArgument,';');
SetControlVisible('BDelete',True);
If Action = 'CREATION' then
begin
   SetControlEnabled('BImprimer',False);
   SetControlEnabled('BDelete',False);
   AnneeAttest := RechDom('PGANNEE',AnneeAttest,False);
  If MoisAttest <> '' then
  begin
    aa := StrToInt(AnneeAttest);
    mm := StrToInt(MoisAttest);
    DateDebutAttest := EncodeDate(aa,mm,1);
    DateFinAttest := FinDeMois(dateDebutAttest);
  end
  else
  begin
    aa := StrToInt(AnneeAttest);
    DateDebutAttest := EncodeDate(aa,1,1);
    DateFinAttest := EncodeDate(aa,12,31);
  end;
end
else
begin
     SetControlEnabled('BTPER',False);
     SetControlEnabled('BPERPREC',False);
     SetControlEnabled('BPERSUIV',False);
     SetControlEnabled('CINCREMENTAUTO',False);
     SetControlText('ENUMDOC',LeNumAEM);
end;
Q := OpenSQL('SELECT PEL_MONTANTEURO FROM ELTNATIONAUX WHERE ##PEL_PREDEFINI## PEL_CODEELT="0095" ' +
    'AND PEL_DATEVALIDITE<="' + UsDateTime(DateFinAttest) + '" ORDER BY PEL_DATEVALIDITE DESC', true);
if not Q.Eof then TauxEltNat := Q.Fields[0].AsFloat
else TauxEltNat := 0;
Ferme(Q);

DefautTH:=THEdit(GetControl('ESALBTRUTSAP'));
If DefautTH<>NIL Then DefautTH.OnChange:=CalculCotisation;
DefautTH:=THEdit(GetControl('ETAUX'));
If DefautTH<>NIL Then DefautTH.OnChange:=CalculCotisation;
//PT26 Ajout ===>
DefautTH:=THEdit(GetControl('ESALBTRUTSAP1'));
If DefautTH<>NIL Then DefautTH.OnChange:=CalculCotisation;
DefautTH:=THEdit(GetControl('ETAUX1'));
If DefautTH<>NIL Then DefautTH.OnChange:=CalculCotisation;
DefautTH:=THEdit(GetControl('ECOTISATION'));
If DefautTH<>NIL Then DefautTH.OnChange:=CalculCotisation;
DefautTH:=THEdit(GetControl('ECOTISATION1'));
If DefautTH<>NIL Then DefautTH.OnChange:=CalculCotisation;
DefautTH:=THEdit(GetControl('HEURESART'));  //Devient Nb de cachets isolés
If DefautTH<>NIL Then DefautTH.OnChange:=GererCachets;
DefautTH:=THEdit(GetControl('CACHETS'));    //Devient Nb de cachets groupés
If DefautTH<>NIL Then DefautTH.OnChange:=GererCachets;
DefautTH:=THEdit(GetControl('NBJOURSTRAV'));
If DefautTH<>NIL Then DefautTH.OnExit:=GererCachets;
//PT26 Fin Ajout <===

DateEdition:=ThEdit(getcontrol('EDATE'));
If DateEdition<>nil then DateEdition.OnElipsisClick:=DateElipsisclick;
BImprimer:=TToolBarButton97(GetControl('BIMPRIMER'));
If BImprimer<>NIL Then BImprimer.OnClick:=impression;

Edit := THEdit(GetControl('DECLARANT'));
If Edit <> Nil then Edit.OnExit := AfficheDeclarant;
Check := TCheckBox(GetControl('CATTESTINIT'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('ATTESTCOMPL'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CRECAPPOS'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CRECAPNEG'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CSALCADRE'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CSALNONCADRE'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CREALISATEUR'));      //PT26
If Check <> Nil then Check.OnClick := GereCheckBox;  //PT26
Check := TCheckBox(GetControl('CARTISTE'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('COUVRIER'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CTECHNICIEN'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CLICENCE'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CLICENCENON'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CORGANISATEUR'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CORGANISATEURNON'));
If Check <> Nil then Check.OnClick := GereCheckBox;;
Check := TCheckBox(GetControl('CLABEL'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CLABELNON'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CAFFILCP'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CAFFILCPNON'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('PARENTE'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('PARENTENON'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CFINCDD'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CRUPTEMPL'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CRUPTSAL'));
If Check <> Nil then Check.OnClick := GereCheckBox;
Check := TCheckBox(GetControl('CENCOURS'));          //PT26
If Check <> Nil then Check.OnClick := GereCheckBox;  //PT26
// Gestion du navigateur

PerPrec:=TToolBarButton97(GetControl('BPERPREC'));         // Navigation si,plus de 4 contrats
If PerPrec<>Nil Then PerPrec.OnClick:=PeriodePrecedente;
PerSuiv:=TToolBarButton97(GetControl('BPERSUIV'));
If PerSuiv<>Nil Then PerSuiv.OnClick:=PeriodeSuivante;
MenuPeriodes := TPopUpMenu(GetControl('TPPERIODES'));
AfficherSalarie;
Edit := THEdit(GetControl('ANNEE'));
If Edit <> nil then Edit.OnExit := ChangeMoisAnnee;
Combo := THValComboBox(GetControl('MOIS'));
If Combo <> Nil then Combo.OnChange := ChangeMoisAnnee;
Edit := THEdit(GetControl('EDATEDEB'));
If Edit <> nil then Edit.OnExit := ChangeMoisAnnee;
Edit := THEdit(GetControl('EDATEFIN'));
If Edit <> nil then Edit.OnExit := ChangeMoisAnnee;
//DEBUT PT12
Check := TCheckBox(GetControl('CINCREMENTAUTO'));
If Check <> Nil then
begin
     Check.OnClick := AccesNumAttest;
     Check.Checked := True;
end;
//FIN PT12
BVal := TToolBarButton97(GetControl('BValider'));
If BVal <> Nil then BVal.OnClick := ValidationAttest;
BSupp := TToolBarButton97(GetControl('BDelete'));
If BSupp <> Nil then BSupp.OnClick := SuppAttest;
SetControlProperty('BValider','ShowHint',True);
SetControlProperty('BDelete','ShowHint',True);
//DEBUT PT22
SetControlEnabled('CAFFILCP',False);
SetControlEnabled('ENUMAFFILIATION',False);
SetControlEnabled('CAFFILCPNON',False);
//FIN PT22
end;

procedure TOF_PGASSEDIC_SPECTACLE.ChargeZonesEmployeur;
var Q:TQuery;
    Cinter : String; //PT26
begin
//PT26 Ajout champ ET_EMAIL dans réquete SQL ==
Q:=OpenSQL('SELECT ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_JURIDIQUE,ET_ACTIVITE,ET_TELEPHONE,ET_FAX,ET_SIRET,ET_APE,'+
        'ET_CODEPOSTAL,ET_VILLE,ET_EMAIL,ETB_ETABLISSEMENT,ETB_ISNUMRECOUV,ETB_ISLICSPEC,ETB_ISNUMLIC,ETB_ISOCCAS,ETB_ISLABELP,'+
        'ETB_ISNUMLAB,ETB_ISNUMCPAY FROM ETABCOMPL '+
        'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT'+
                   ' WHERE PSA_SALARIE="'+LeSalarie+'"',True);
If Not Q.eof then        //PortageCWAS
   begin
   SetControlText('ESOCIETE',Q.FindField('ET_LIBELLE').AsString);
{PT27
   SetControlText('EADRETAB1',Q.FindField('ET_ADRESSE1').AsString);
   SetControlText('EADRETAB12',Q.FindField('ET_ADRESSE2').AsString + Q.FindField('ET_ADRESSE3').AsString);
   SetControlText('EFORMJURIDIQUE',RechDom('TTFORMEJURIDIQUE',Q.FindField('ET_JURIDIQUE').AsString,False));  //PT9
   SetControlText('EACTIVITEETAB',Q.FindField('ET_ACTIVITE').AsString);
}

   //PT26 SetControlText('ETELETAB',Q.FindField('ET_TELEPHONE').AsString);
   Cinter :='';                                                  //PT26
   ForceNumerique(Q.FindField('ET_TELEPHONE').AsString,Cinter);  //PT26
   SetControlText('ETELETAB',Cinter);                            //PT26
   //PT26 SetControlText('EFAXETAB',Q.FindField('ET_FAX').AsString);
   Cinter :='';                                                  //PT26
   ForceNumerique(Q.FindField('ET_FAX').AsString,Cinter);        //PT26
   SetControlText('EFAXETAB',Cinter);                            //PT26

   SetControlText('ESIRET',Q.FindField('ET_SIRET').AsString);
   SetControlText('EAPE',Q.FindField('ET_APE').AsString);
   SetControlText('ECPETAB',Q.FindField('ET_CODEPOSTAL').AsString);
   SetControlText('ECOMMUNE',Q.FindField('ET_VILLE').AsString);
   SetControlText('ENUMCENTREREC',Q.FindField('ETB_ISNUMRECOUV').AsString);
   If Q.FindField('ETB_ISLICSPEC').AsString='X' Then
      begin
              SetControlChecked('CLICENCE',True);
              SetControlText('ENUMLICENCE',Q.FindField('ETB_ISNUMLIC').AsString);
      end
   Else
      begin
              SetControlChecked('CLICENCENON',True);
      end;
   If Q.FindField('ETB_ISLABELP').AsString='X' Then
      begin
              SetControlChecked('CLABEL',True);
              SetControlText('ENUMLABEL',Q.FindField('ETB_ISNUMLAB').AsString);
      end
   Else
      begin
              SetControlChecked('CLABELNON',True);
      end;
   If Q.FindField('ETB_ISOCCAS').AsString='X' Then SetControlChecked('CORGANISATEUR',True)
   Else SetControlChecked('CORGANISATEURNON',true);
   if (Q.FindField('ETB_ISNUMCPAY').AsString) <> '' then
   begin
           SetControltext('ENUMAFFILIATION',Q.FindField('ETB_ISNUMCPAY').AsString);
           SetControlChecked('CAFFILCP',True);
   end
   else SetControlChecked('CAFFILCPNON',True);
   end;
   SetControlText('EMAILETAB',Q.FindField('ET_EMAIL').AsString);  //PT26
Ferme(Q);
end ;

procedure TOF_PGASSEDIC_SPECTACLE.CalculCotisation(Sender:TObject) ;
Var Montant:Double;
    TEdit : THEdit;  //PT26
    Mnt   : Double;  //PT26
begin
  //PT26 ======>
  TEdit := THEdit(Sender);
  If TEdit = nil Then Exit;
  If (Tedit.name = 'ESALBTRUTSAP') Or (Tedit.name = 'ETAUX') Then
  Begin   //PT26 <======
     If (IsNumeric(GetControlText('ETAUX'))) AND (IsNumeric(GetControlText('ESALBTRUTSAP'))) Then
     begin
       Montant:=(StrToFloat(GetControltext('ETAUX')))*(StrToFloat(GetControlText('ESALBTRUTSAP')))/100;
       SetControltext('ECOTISATION',FloatToStr(Arrondi(Montant,0)));
     end;
  End; //PT26
  //PT26 ======>
  If Tedit.name = 'ESALBTRUTSAP1' Then
  Begin
     Montant := 0;
     Mnt     := 0;
     If IsNumeric(GetControlText('ESALBTRUTSAP1')) Then Montant := StrToFloat(GetControltext('ESALBTRUTSAP1'));
     If IsNumeric(GetControlText('ETAUX1')) Then Mnt := StrToFloat(GetControltext('ETAUX1'));
     If (Montant <> 0) And (Mnt = 0) Then SetControlText('ETAUX1',GetControltext('ETAUX'));
  End;

  If (Tedit.name = 'ESALBTRUTSAP1') Or (Tedit.name = 'ETAUX1') Then
  Begin
     If (IsNumeric(GetControlText('ETAUX1'))) AND (IsNumeric(GetControlText('ESALBTRUTSAP1'))) Then
     begin
       Montant:=(StrToFloat(GetControltext('ETAUX1')))*(StrToFloat(GetControlText('ESALBTRUTSAP1')))/100;
       SetControltext('ECOTISATION1',FloatToStr(Arrondi(Montant,0)));
     end;
  End;
  If (IsNumeric(GetControlText('ECOTISATION'))) AND (IsNumeric(GetControlText('ECOTISATION1'))) Then
  begin
     Montant:=(StrToFloat(GetControltext('ECOTISATION')))+(StrToFloat(GetControlText('ECOTISATION1')));
     SetControltext('ECOTISTOTAL',FloatToStr(Montant));
  End;
  //PT26 <======
end;

//PT26 Ajout procedure GererCachets ============================
procedure TOF_PGASSEDIC_SPECTACLE.GererCachets(Sender:TObject) ;
Var TEdit : THEdit;

Begin
  TEdit := THEdit(Sender);
  If TEdit = nil Then Exit;

  If Tedit.name = 'HEURESART' Then
  Begin
     If IsNumeric(GetControlText('HEURESART')) Then
     Begin
        If StrToFloat(GetControltext('HEURESART')) <> 0 Then SetControltext('CACHETS','0');
     End;
  End;
  If Tedit.name = 'CACHETS' Then
  Begin
     If IsNumeric(GetControlText('CACHETS')) Then
     Begin
        If StrToFloat(GetControltext('CACHETS')) <> 0 Then SetControltext('HEURESART','0');
     End;
  End;
End;

//PT26 Ajout Function CtrlSaisie(Ctrl : Boolean)========
//PT28 Function TOF_PGASSEDIC_SPECTACLE.CtrlSaisie : Boolean;
Function TOF_PGASSEDIC_SPECTACLE.CtrlSaisie(MajBase : Boolean) : Boolean;
Var MESS : array[1..12] of String;   //PT28 array[1..10]
    i :Integer;
    Erreur : String;
    AFaire : Boolean;
Begin
  For i := 1 To 12 do   //PT28 i := 1 To 10 do
  Begin
    MESS[i] := '';
  End;
  i := 0;
  AFaire := True;
  If Not IsNumeric(GetControlText('HEURESOUV')) Then SetControltext('HEURESOUV','0');
  If Not IsNumeric(GetControlText('HEURESART')) Then SetControltext('HEURESART','0');
  If Not IsNumeric(GetControlText('CACHETS')) Then SetControltext('CACHETS','0');
  If Not IsNumeric(GetControlText('NBJOURSTRAV')) Then SetControltext('NBJOURSTRAV','0');
  If Not IsNumeric(GetControlText('ESALBTRUTSAV')) Then SetControltext('ESALBTRUTSAV','0');
  If Not IsNumeric(GetControlText('ESALBTRUTSAP')) Then SetControltext('ESALBTRUTSAP','0');
  If Not IsNumeric(GetControlText('ETAUX')) Then SetControltext('ETAUX','0');
  If Not IsNumeric(GetControlText('ECOTISATION')) Then SetControltext('ECOTISATION','0');
  If Not IsNumeric(GetControlText('ESALBTRUTSAP1')) Then SetControltext('ESALBTRUTSAP1','0');
  If Not IsNumeric(GetControlText('ETAUX1')) Then SetControltext('ETAUX1','0');
  If Not IsNumeric(GetControlText('ECOTISATION1')) Then SetControltext('ECOTISATION1','0');

  If (GetCheckBoxState('CATTESTINIT') = CbUnChecked) And (GetCheckBoxState('ATTESTCOMPL') = CbUnChecked) And
     (GetCheckBoxState('CRECAPNEG') = CbUnChecked) And (GetCheckBoxState('CRECAPPOS') = CbUnChecked) then
  Begin
     i := i + 1;
     MESS[i] := '* Le Type d''AEM est doit être renseigné';
     SetFocusControl('CATTESTINIT');
  End;
  If ((GetCheckBoxState('ATTESTCOMPL') = CbChecked) Or (GetCheckBoxState('CRECAPNEG') = CbChecked) Or
     (GetCheckBoxState('CRECAPPOS') = CbChecked)) And (Trim(GetControlText('NUMATTESTINIT'))='') then
  Begin
     i := i + 1;
     MESS[i] := '* Le Numéro attestation initiale doit être renseigné';
     SetFocusControl('NUMATTESTINIT');
  End;
  If (GetCheckBoxState('CARTISTE') = CbUnChecked) And (GetCheckBoxState('COUVRIER') = CbUnChecked) And
     (GetCheckBoxState('CTECHNICIEN') = CbUnChecked) And (GetCheckBoxState('CREALISATEUR') = CbUnChecked) then
  Begin
     i := i + 1;
     MESS[i] := '* La catégorie de l''intermittent doit être renseigné';
     SetFocusControl('CREALISATEUR');
  End;
  If (GetControlText('EDATEFIN') <> DateToStr(IDate1900)) And (GetCheckBoxState('CFINCDD') = CbUnChecked) And
     (GetCheckBoxState('CRUPTEMPL') = CbUnChecked) And (GetCheckBoxState('CRUPTSAL') = CbUnChecked) then
  Begin
     i := i + 1;
     MESS[i] := '* Le motif de cessation du contrat de travail doit être renseigné';
     SetFocusControl('CFINCDD');
  End;
  If (GetCheckBoxState('CATTESTINIT') = CbChecked) Or (GetCheckBoxState('ATTESTCOMPL') = CbChecked) Then
  Begin
     If (StrToFloat(GetControltext('ESALBTRUTSAP')) <> 0) Or (StrToFloat(GetControltext('ECOTISATION')) <> 0) Then
     Begin
        If (StrToFloat(GetControltext('HEURESOUV'))=0) And
           ((GetCheckBoxState('CTECHNICIEN') = CbChecked) Or (GetCheckBoxState('COUVRIER') = CbChecked)) Then
        Begin
           i := i + 1;
           MESS[i] := '* Le nombre d''heures effectuées doit être déclaré';
           SetFocusControl('HEURESOUV');
        End;
        If (StrToFloat(GetControltext('HEURESOUV'))=0) And (StrToFloat(GetControltext('HEURESART')) = 0) And (StrToFloat(GetControltext('CACHETS')) = 0) And
           ((GetCheckBoxState('CARTISTE') = CbChecked) Or (GetCheckBoxState('CREALISATEUR') = CbChecked)) Then
        Begin
           i := i + 1;
           MESS[i] := '* Le nombre d''heures effectuées et/ou le nombre de cachets doivent être déclarés';
           SetFocusControl('HEURESOUV');
        End;
        If StrToFloat(GetControltext('NBJOURSTRAV'))=0 Then
        Begin
           i := i + 1;
           MESS[i] := '* Le nombre de jours travaillés doit être déclaré';
           SetFocusControl('NBJOURSTRAV');
           AFaire := False;
        End;
     End;
     If (StrToFloat(GetControltext('HEURESART')) <> 0) Or (StrToFloat(GetControltext('CACHETS')) <> 0) Then
     Begin
        If (StrToFloat(GetControltext('NBJOURSTRAV'))=0) And (AFaire = True) Then
        Begin
           i := i + 1;
           MESS[i] := '* Le nombre de jours travaillés doit être déclaré';
           SetFocusControl('NBJOURSTRAV');
        End;
        If StrToFloat(GetControltext('NBJOURSTRAV')) <> 0 Then
        Begin
           If StrToFloat(GetControltext('NBJOURSTRAV')) < 5 Then
           Begin
              If (StrToFloat(GetControltext('HEURESART')) = 0) And (StrToFloat(GetControltext('CACHETS')) <> 0) Then
              Begin
                i := i + 1;
                MESS[i] := '* Le nombre de jours travaillés est inférieur à 5 jours, le nombre de cachets isolés a été ventilé';
                SetControltext('HEURESART',GetControltext('CACHETS'));
                SetControltext('CACHETS','0');
                SetFocusControl('HEURESART');
              End;
           End
           Else
           Begin
              If (StrToFloat(GetControltext('CACHETS')) = 0) And (StrToFloat(GetControltext('HEURESART')) <> 0) Then
              Begin
                i := i + 1;
                MESS[i] := '* Le nombre de jours travaillés est supérieur ou égal à 5 jours, le nombre de cachets groupés a été ventilé';
                SetControltext('CACHETS',GetControltext('HEURESART'));
                SetControltext('HEURESART','0');
                SetFocusControl('CACHETS');
              End;
           End;
        End;
     End;
     If (StrToFloat(GetControltext('ESALBTRUTSAP1')) = 0) And (StrToFloat(GetControltext('ECOTISATION1')) = 0) Then
     Begin
        If (StrToFloat(GetControltext('ESALBTRUTSAP')) = 0) Or (StrToFloat(GetControltext('ETAUX')) = 0) Or
           (StrToFloat(GetControltext('ECOTISATION')) = 0) Or (StrToFloat(GetControltext('ESALBTRUTSAV')) = 0) Then
        Begin
           i := i + 1;
           MESS[i] := '* La rémunération principale est obligatoire si Autres rémunérations non renseignées';
           SetFocusControl('ESALBTRUTSAV');
        End;
     End;
  End;
  //PT28 Debut Ajout ==>
  If Trim(GetControlText('NUMOBJET'))='' Then
  Begin
     i := i + 1;
     MESS[i] := '* Le N° d''objet n''est pas saisi';
     SetFocusControl('NUMOBJET');
  End;
  //PT28 Fin Ajout <==
  If i > 0 Then
  Begin
     Erreur := '';
     For i := 1 To 12 do   //PT28 i := 1 To 10 do
     Begin
        If MESS[i] = '' Then Break;
        Erreur := Erreur + MESS[i];
        //PT28 If MESS[i+1] <> '' THEN Erreur := Erreur + chr(13)+ ' ';
        Erreur := Erreur + chr(13)+ ' '; //PT28
     End;
     //PT28 Debut Ajout ==>
     If MajBase = FALSE Then Erreur := Erreur + 'Voulez-vous continuer l''impression?'
                        Else Erreur := Erreur + 'Voulez-vous enregistrer la saisie?';

     If PGIASK(Erreur,''+TFVierge(Ecran).Caption+'') = mrYes Then Result := True
                                                             Else Result := False;
      //PT28 Fin Ajout <==
     //PT28 PGIBOX(Erreur,''+TFVierge(Ecran).Caption+'');
     //PT28 Result := False;
  End
  Else Result := True;
End;

procedure TOF_PGASSEDIC_SPECTACLE.impression(Sender:TObject) ;   // Lancement de l'état
var //PT26 Q : TQuery;
    //PT26 TobEtab,TobEdition,TEt,T : Tob;
  TobEdition,T : Tob;  //PT26
  //PT26 Salarie,Etab,NumDoc ,Categ,St: String;
  Salarie,NumDoc,St: String;  //PT26
  //PT26 i : Integer;
  Pages : TPageControl;
  //PT26 ImpSal,ImpAssedic,ImpDupli : Boolean;
  //PT26 EtatsAImprimer,Etat : String;
  //PT26 mm,yy : word;
  //PT26 DatePeriodeEC : TDateTime;
  //PT26 An : String;
  DateNaiss,DateAttest,DateD,DateF : TDateTime;
begin
  Pages := TPageControl(GetControl('PAGES'));
  TobEdition := Tob.Create('Edition',Nil,-1);
  //PT28 CtrlSaisie;  //PT26
  If CtrlSaisie(FALSE) = False Then Exit;   //PT28

  NumDoc := GetControlText('ENUMDOC');
  T := Tob.Create('LigneEdition',TobEdition,-1);
  //PT26 ajout du UpperCase lignes ci-dessous **
  T.AddChampSupValeur('PSA_SALARIE',Salarie);
  T.AddChampSupValeur('ENUMDOC',UpperCase(NumDoc));
  T.AddChampSupValeur('PGA_NUMATTEST',UpperCase(NumDoc));
  T.AddChampSupValeur('LNOM',UpperCase(GetControlText('LNOM')));
  T.AddChampSupValeur('CENCOURS',GetControlText('CENCOURS'));
{PT27
  T.AddChampSupValeur('PSA_NOMJF',UpperCase(GetControlText('PSA_NOMJF')));
}
  T.AddChampSupValeur('LPRENOM',UpperCase(GetControlText('LPRENOM')));
  T.AddChampSupValeur('CATTESTINIT','-');
  T.AddChampSupValeur('ATTESTCOMPL','-');
  T.AddChampSupValeur('CRECAPPOS','-');
  T.AddChampSupValeur('CRECAPNEG','-');   //PT26 RECAPNEG=>CRECAPNEG

  iF GetCheckBoxState('ATTESTCOMPL') = CbChecked then T.PutValue('ATTESTCOMPL','X')
  else iF GetCheckBoxState('CRECAPNEG') = CbChecked then T.PutValue('CRECAPNEG','X')
  else iF GetCheckBoxState('CRECAPPOS') = CbChecked then T.PutValue('CRECAPPOS','X')
  else T.PutValue('CATTESTINIT','X');

  T.AddChampSupValeur('EEMPLOI',UpperCase(GetControlText('EEMPLOI')));
  T.AddChampSupValeur('ERETRAITE',UpperCase(GetControlText('ERETRAITE')));
{PT27
  T.AddChampSupValeur('ENUMASSEDIC',GetControlText('ENUMASSEDIC'));
}
  //PT26 T.AddChampSupValeur('PSA_NUMEROSS',GetControlText('ENUMIDENTIFIANT'));
  T.AddChampSupValeur('ENUMIDENTIFIANT',GetControlText('ENUMIDENTIFIANT'));
  //DEBUT PT23
  If IsValidDate(GetControlText('EDATENAISSANCE')) then DateNaiss := StrToDate(GetControlText('EDATENAISSANCE'))
  else DateNaiss := IDate1900;
  T.AddChampSupValeur('EDATENAISSANCE',DateNaiss);
  //FIN PT23
{PT27
  T.AddChampSupValeur('ETELSALARIE',GetControlText('ETELSALARIE'));
}
  //PT26 T.AddChampSupValeur('PSA_CODEPOSTAL',GetControlText('ECPSAL'));
  T.AddChampSupValeur('ECPSAL',GetControlText('ECPSAL'));
  //PT26 T.AddChampSupValeur('PSA_VILLE',GetControlText('ECOMMUNESAL'));
  T.AddChampSupValeur('ECOMMUNESAL',UpperCase(GetControlText('ECOMMUNESAL')));
  T.AddChampSupValeur('EADRSAL1',UpperCase(GetControlText('EADRSAL1')));
{PT27
  T.AddChampSupValeur('EADRSAL2',UpperCase(GetControlText('EADRSAL2')) + UpperCase(GetControlText('EADRSAL3')));
}
  T.AddChampSupValeur('EADRSAL2',UpperCase(GetControlText('EADRSAL2')));
//FIN PT27
  //PT26 T.AddChampSupValeur('ENOMUSAGE',GetControlText('PSA_SURNOM'));
  T.AddChampSupValeur('ENOMUSAGE',UpperCase(GetControlText('ENOMUSAGE')));
  //PT26 T.AddChampSupValeur('MOIS',GetControlText('PGA_MOIS'));
  T.AddChampSupValeur('MOIS',GetControlText('MOIS'));
  T.AddChampSupValeur('ANNEE',RechDom('PGANNEE',GetControlText('ANNEE'),False));
  T.AddChampSupValeur('NUMATTESTINIT',UpperCase(GetControlText('NUMATTESTINIT')));

  If GetCheckBoxState('CSALCADRE')= CbChecked then
  begin
       T.AddChampSupValeur('CSALCADRE','X');
       T.AddChampSupValeur('CSALNONCADRE','-');
  end
  else
  begin
       T.AddChampSupValeur('CSALCADRE','-');
       T.AddChampSupValeur('CSALNONCADRE','X');
  end;
  //DEBUT PT23
  If IsValidDate(GetControlText('EDATEDEB')) then DateD := StrToDate(GetControlText('EDATEDEB'))
  else DateD := IDate1900;
  T.AddChampSupValeur('EDATEDEB',DateD);
  If IsValidDate(GetControlText('EDATEFIN')) then DateF := StrToDate(GetControlText('EDATEFIN'))
  else DateF := IDate1900;
  T.AddChampSupValeur('EDATEFIN',DateF);
  If IsValidDate(GetControlText('EDATE')) then DateAttest := StrToDate(GetControlText('EDATE'))
  else DateAttest := IDate1900;
  T.AddChampSupValeur('EDATE',DateAttest);
  //FIN PT23
  T.AddChampSupValeur('LIENPARENTE',UpperCase(GetControlText('LIENPARENTE')));
  //PT26 T.AddChampSupValeur('LELIENPARENTE',GetControlText('LIENPARENTE'));
  //PT26 T.AddChampSupValeur('PARENTE',GetControlText('PARENTE'));
  //PT26 Debut ==>
  If GetCheckBoxState('PARENTE')= CbChecked then
  begin
       T.AddChampSupValeur('PARENTE','X');
       T.AddChampSupValeur('PARENTENON','-');
  end
  else
  begin
       T.AddChampSupValeur('PARENTE','-');
       T.AddChampSupValeur('PARENTENON','X');
  end;
  //PT26 <== Fin
  T.AddChampSupValeur('HEURESOUV',StrToFloat(GetControlText('HEURESOUV')));
  T.AddChampSupValeur('CREALISATEUR',GetControlText('CREALISATEUR')); //PT26
  T.AddChampSupValeur('CARTISTE',GetControlText('CARTISTE'));
  T.AddChampSupValeur('CTECHNICIEN',GetControlText('CTECHNICIEN'));
  T.AddChampSupValeur('COUVRIER',GetControlText('COUVRIER'));
  T.AddChampSupValeur('HEURESART',StrToFloat(GetControlText('HEURESART')));
  T.AddChampSupValeur('CACHETS',StrToFloat(GetControlText('CACHETS')));
  T.AddChampSupValeur('NBJOURSTRAV',StrToFloat(GetControlText('NBJOURSTRAV')));
  T.AddChampSupValeur('ESALBTRUTSAV',StrToFloat(GetControlText('ESALBTRUTSAV')));
  T.AddChampSupValeur('ESALBTRUTSAP',StrToFloat(GetControlText('ESALBTRUTSAP')));
  T.AddChampSupValeur('ETAUX',StrToFloat(GetControlText('ETAUX')));
  T.AddChampSupValeur('ECOTISATION',StrToFloat(GetControlText('ECOTISATION')));
  //PT26 Debut ==>
  T.AddChampSupValeur('ESALBTRUTSAP1',StrToFloat(GetControlText('ESALBTRUTSAP1')));
  T.AddChampSupValeur('ETAUX1',StrToFloat(GetControlText('ETAUX1')));
  T.AddChampSupValeur('ECOTISATION1',StrToFloat(GetControlText('ECOTISATION1')));
  T.AddChampSupValeur('ECOTISTOTAL',StrToFloat(GetControlText('ECOTISTOTAL')));
  T.AddChampSupValeur('NUMOBJET',UpperCase(GetControlText('NUMOBJET')));
  //PT26 <== Fin
  T.AddChampSupValeur('ESOCIETE',UpperCase(GetControlText('ESOCIETE')));
  //PT26 T.AddChampSupValeur('EADRETAB1',GetControlText('EADRETAB1'));
  //PT26 T.AddChampSupValeur('EADRETAB2',GetControlText('EADRETAB2') + GetControlText('EADRETAB3'));
  //PT26 T.AddChampSupValeur('JURIDIQUE',RechDom('TTFORMEJURIDIQUE',GetControlText('JURIDIQUE'),False));  //PT9
  //PT26 T.AddChampSupValeur('ET_ACTIVITE',GetControlText('EACTIVITEETAB'));
  //PT26 T.AddChampSupValeur('ET_TELEPHONE',GetControlText('ETELETAB'));
  T.AddChampSupValeur('ETELETAB',GetControlText('ETELETAB'));
  T.AddChampSupValeur('EFAXETAB',GetControlText('EFAXETAB'));
  T.AddChampSupValeur('ESIRET',GetControlText('ESIRET'));
  T.AddChampSupValeur('EAPE',UpperCase(GetControlText('EAPE')));
  //PT26 T.AddChampSupValeur('CODEPOSTAL',GetControlText('ECPETAB'));
  T.AddChampSupValeur('ECPETAB',GetControlText('ECPETAB'));
  //PT26 T.AddChampSupValeur('ET_VILLE',GetControlText('ECOMMUNE'));
  T.AddChampSupValeur('ECOMMUNE',UpperCase(GetControlText('ECOMMUNE')));
  T.AddChampSupValeur('ENUMCENTREREC',UpperCase(GetControlText('ENUMCENTREREC')));
  T.AddChampSupValeur('EMAILETAB',UpperCase(GetControlText('EMAILETAB')));
  T.AddChampSupValeur('CLICENCE',GetControlText('CLICENCE'));
  //PT26 T.AddChampSupValeur('ENUMLICENCE',GetControlText('ETB_ISNUMLIC'));
  T.AddChampSupValeur('ENUMLICENCE',UpperCase(GetControlText('ENUMLICENCE')));
  T.AddChampSupValeur('CLABEL',GetControlText('CLABEL'));
  T.AddChampSupValeur('CLABELNON',GetControlText('CLABELNON'));
  //PT26 T.AddChampSupValeur('ENUMLABEL',GetControlText('ETB_ISNUMLAB'));
  T.AddChampSupValeur('ENUMLABEL',UpperCase(GetControlText('ENUMLABEL')));
  T.AddChampSupValeur('CORGANISATEUR',GetControlText('CORGANISATEUR'));
  T.AddChampSupValeur('CORGANISATEURNON',GetControlText('CORGANISATEURNON'));
  T.AddChampSupValeur('ENUMAFFILIATION',UpperCase(GetControlText('ENUMAFFILIATION')));
  T.AddChampSupValeur('CAFFILCP',GetControlText('CAFFILCP'));
  T.AddChampSupValeur('CAFFILCPNON',GetControlText('CAFFILCPNON'));

  St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
  T.AddChampSupValeur('ENOMEMPL', UpperCase(GetControlText('ENOMEMPL')));
  T.AddChampSupValeur('EPRENOMEMPL',UpperCase(GetControlText('EPRENOMEMPL')));
  T.AddChampSupValeur('ELIEU', UpperCase(GetControlText('ELIEU')));
  T.AddChampSupValeur('PERSAJOINDRE',UpperCase(GetControlText('PERSAJOINDRE')));
  T.AddChampSupValeur('TELPERSAJOINDRE', GetControlText('TELPERSAJOINDRE'));
  //DEBUT PT16
  T.AddChampSupValeur('CFINCDD', GetControlText('CFINCDD'));
  T.AddChampSupValeur('CRUPTEMPL', GetControlText('CRUPTEMPL'));
  T.AddChampSupValeur('CRUPTSAL', GetControlText('CRUPTSAL'));
  //FIN PT16
  St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
  if St = 'AUT' then T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False)))
  else T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGQUALDECLARANT2', St, False)));

  ExecuteSQL('UPDATE PGAEM SET PGA_FAIT="X" WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"');
  LanceEtatTOB('E','PAT','AS1',TobEdition,True,False,False,Pages,'','',False,0,'');
  LanceEtatTOB('E','PAT','AS2',TobEdition,True,False,False,Pages,'','',False,0,'');
  LanceEtatTOB('E','PAT','AS2',TobEdition,True,False,False,Pages,'','',True,0,'');
  TobEdition.Free;
end;

procedure TOF_PGASSEDIC_SPECTACLE.DateElipsisClick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGASSEDIC_SPECTACLE.MetABlanc(ZoneSal,ZonePeriode : Boolean);
begin
If ZoneSal = True then
begin
        SetControlText('LMATRICULE','');
        SetControlText('LNOM','');
        SetControlText('LPRENOM','');
{PT27
        SetControlText('ENUMASSEDIC','');
}
        SetControlText('ENUMIDENTIFIANT','');
        SetControltext('ERETRAITE','');
        SetControlText('EEMPLOI','');
        SetControlText('ESALBTRUTSAV','');
        SetControltext('ESALBTRUTSAP','');
        SetControlText('ETAUX','');
        SetControlText('ECOTISATION','');
        SetControlText('ESALBTRUTSAP1','');  //PT26
        SetControlText('ETAUX1','');         //PT26
        SetControlText('ECOTISATION1','');   //PT26
        SetControlText('ECOTISTOTAL','');    //PT26
        SetControlText('ENUMLICENCE','');
        SetControlText('ENUMLABEL','');
        SetControltext('ENUMAFFILIATION','');
        SetControlText('ENUMCENTREREC','');
end;
If ZonePeriode = True then
begin
        SetControlText('HEURESOUV','0');
        SetControlText('HEURESART','0');
        SetControlText('NBJOURSTRAV','0');
        SetControlText('CACHETS','0');
        SetControlText('ESALBTRUTSAV','0');
        SetControlText('ESALBTRUTSAP','0');
        SetControlText('ETAUX','0');
        SetControlText('ECOTISATION','0');
        SetControlText('ESALBTRUTSAP1','0');  //PT26
        SetControlText('ETAUX1','0');         //PT26
        SetControlText('ECOTISATION1','0');   //PT26
        SetControlText('ECOTISTOTAL','0');    //PT26
        SetControlText('EDATEDEB',DateToStr(IDate1900));
        SetControlText('EDATEFIN',DateToStr(IDate1900));
end;
end;

Procedure TOF_PGASSEDIC_SPECTACLE.RemplirTobLesPeriodes;
var TobTemp,TLP : Tob;
    x,i : Integer;
    Q : TQuery;
    St,Origine,Emploi : String;
    NewItems : TMenuItem;
    BPeriode : TToolBarButton97;
    DD,DF,DFc,DDc : TDateTime;
    BrutAv,BrutAp,Jours : Double;
    RubJours,RubCachets : String;
    Typecal,TypeBase,TypeMontant,ChampBul : String;
    TypecalC,TypeBaseC,TypeMontantC,ChampBulC : String;
    NumObjet : String; //PT28
begin
     //DEBUT PT13
        RubJours := GetParamSoc('SO_PGASSSPECRUBJOURS');
        Q := OpenSQL('SELECT PRM_CODECALCUL,PRM_TYPEBASE,PRM_TYPEMONTANT FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="'+RubJours+'"',True);
        If Not Q.Eof then
        begin
             Typecal := Q.FindField('PRM_CODECALCUL').AsString;
             TypeBase := Q.FindField('PRM_TYPEBASE').AsString;
             TypeMontant := Q.FindField('PRM_TYPEMONTANT').AsString;
        end;
        Ferme(Q);
        If TypeCal = '01' then
        begin
             If typeBase = '00' then ChampBul := 'PHB_BASEREM'
             else ChampBul := 'PHB_MTREM';
        end
        else ChampBul := 'PHB_BASEREM';
        //FIN PT13
        //DEBUT PT25
        RubCachets := GetParamSocSecur('SO_PGASSSPECRUBCACH','0006'); //PT25
        Q := OpenSQL('SELECT PRM_CODECALCUL,PRM_TYPEBASE,PRM_TYPEMONTANT FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="'+RubCachets+'"',True);
        If Not Q.Eof then
        begin
             TypecalC := Q.FindField('PRM_CODECALCUL').AsString;
             TypeBaseC := Q.FindField('PRM_TYPEBASE').AsString;
             TypeMontantC := Q.FindField('PRM_TYPEMONTANT').AsString;
        end;
        Ferme(Q);
        If TypeCalC = '01' then
        begin
             If typeBaseC = '00' then ChampBulC := 'PHB_BASEREM'
             else ChampBulC := 'PHB_MTREM';
        end
        else ChampBulC := 'PHB_BASEREM';
        //FIN PT25
        TobLesPeriodes := Tob.Create('LesPeriodes',Nil,-1);
        //PT28 Ajout champ PCI_NUMOBJETAEM dans le select
        Q := OpenSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN,PPU_CBRUT,PPU_CHEURESTRAV,'+
           'PCI_ISHEURES,PCI_ISCACHETS,PCI_ISNBEFFECTUE,PCI_MOTIFSORTIE,PCI_LIBELLEEMPLOI,PCI_NUMOBJETAEM FROM PAIEENCOURS'+
           ' LEFT JOIN CONTRATTRAVAIL ON PPU_SALARIE=PCI_SALARIE AND PPU_DATEDEBUT=PCI_DEBUTCONTRAT AND PPU_DATEFIN=PCI_FINCONTRAT WHERE PPU_SALARIE="'+LeSalarie+'"'+
           ' AND PCI_SALARIE IS NOT NULL'+
           ' AND PPU_DATEDEBUT>="'+UsDateTime(DateDebutAttest)+'" AND PPU_DATEDEBUT<="'+UsDateTime(DateFinAttest)+'"'+
           ' AND PPU_DATEFIN>="'+UsDateTime(DateDebutAttest)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFinAttest)+'" ORDER BY PPU_DATEDEBUT',True);
        TobTemp := Tob.Create('Tobtemp',Nil,-1);
        TobTemp.LoadDetailDB('Tobtemp','','',Q,False);
        Ferme(Q);
        For x := 0 to TobTemp.detail.Count - 1 do
        begin
                TLP := Tob.Create('UnePeriode',TobLesPeriodes,-1);
                DD := TobTemp.Detail[x].GetValue('PPU_DATEDEBUT');
                DF := TobTemp.Detail[x].GetValue('PPU_DATEFIN');
                Emploi := TobTemp.Detail[x].GetValue('PCI_LIBELLEEMPLOI'); //PT18
                BrutAv := TobTemp.Detail[x].GetValue('PPU_CBRUT');
                BrutAp := TobTemp.Detail[x].GetValue('PPU_CBRUT');
                NumObjet := TobTemp.Detail[x].GetValue('PCI_NUMOBJETAEM'); //PT28

                Q := OpenSQL('SELECT SUM(PHB_MTREM) MONTANT FROM HISTOBULLETIN '+
                'LEFT JOIN REMUNERATION ON PHB_RUBRIQUE=PRM_RUBRIQUE '+
                'WHERE PHB_SALARIE="'+LeSalarie+'" AND PHB_DATEDEBUT="'+UsDateTime(DD)+'" '+
                'AND PHB_DATEFIN="'+UsdateTime(DF)+'" AND PRM_THEMEREM="ABT"',True);
                If Not Q.Eof then BrutAp := BrutAv - Q.FindField('MONTANT').AsFloat;
                Ferme(Q);
                //DEBUT PT13
                Q := OpenSQL('SELECT SUM(('+ChampBul+')) MONTANT FROM HISTOBULLETIN WHERE PHB_SALARIE="'+LeSalarie+'" '+
                'AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+RubJours+'" '+
                'AND PHB_DATEDEBUT>="'+UsDateTime(DD)+'" '+
                'AND PHB_DATEFIN<="'+UsDateTime(DF)+'"',True);
                If Not Q.Eof then Jours := Q.FindField('MONTANT').AsFloat
                else Jours := 0;
                Ferme(Q);
                //FIN PT13
                TLP.AddChampSupValeur('ORIGINE','CP');
                TLP.AddChampSupValeur('DATEDEBUT',DD);
                TLP.AddChampSupValeur('DATEFIN',DF);
                TLP.AddChampSupValeur('LIBDATEDEBUT',DD);
                TLP.AddChampSupValeur('LIBDATEFIN',DF);
                TLP.AddChampSupValeur('NBCACHETS',0);   //@@@@@@@
                TLP.AddChampSupValeur('NBHEURES',0);
                TLP.AddChampSupValeur('NBJOURS',Jours);
                TLP.AddChampSupValeur('BRUT',BrutAV);
                TLP.AddChampSupValeur('NET',BrutAp);
                TLP.AddChampSupValeur('RUPTURE','FIN');
                TLP.AddChampSupValeur('EMPLOI',Emploi);    //PT18
                TLP.AddChampSupValeur('NOBJET',NumObjet);  //PT28

                If TobTemp.Detail[x].GetValue('PCI_MOTIFSORTIE') = '36' then TLP.PutValue('RUPTURE','EMP');
                If TobTemp.Detail[x].GetValue('PCI_MOTIFSORTIE') = '37' then TLP.PutValue('RUPTURE','SAL');
                If TobTemp.Detail[x].GetValue('PCI_ISHEURES') = 'X' then TLP.Putvalue('NBHEURES',TobTemp.Detail[x].GetValue('PCI_ISNBEFFECTUE'))
                else If TobTemp.Detail[x].GetValue('PCI_ISCACHETS') = 'X' then TLP.Putvalue('NBCACHETS',TobTemp.Detail[x].GetValue('PCI_ISNBEFFECTUE'))
                else
                begin
                        Q := OpenSQL('SELECT '+ChampBulC+' FROM HISTOBULLETIN '+
                        'WHERE PHB_SALARIE="'+LeSalarie+'" AND PHB_DATEDEBUT="'+UsDateTime(DD)+'" '+
                        'AND PHB_DATEFIN="'+UsdateTime(DF)+'" AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+RubCachets+'"',True);
                        If Not Q.Eof then TLP.PutValue('NBCACHETS',Q.FindField(ChampBulc).AsFloat)
                        else TLP.Putvalue('NBHEURES',TobTemp.Detail[x].GetValue('PPU_CHEURESTRAV'));
                        Ferme(Q);
                end;
        end;
        TobTemp.Free;
        Q := OpenSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN,PPU_CBRUT,PPU_CHEURESTRAV,PSA_DATEENTREE,PSA_DATESORTIE FROM PAIEENCOURS LEFT JOIN SALARIES '+
        'ON PPU_SALARIE=PSA_SALARIE WHERE PPU_SALARIE="'+LeSalarie+'"'+
           ' AND PPU_DATEDEBUT>="'+UsDateTime(DateDebutAttest)+'" AND PPU_DATEDEBUT<="'+UsDateTime(DateFinAttest)+'"'+
           ' AND PPU_DATEFIN>="'+UsDateTime(DateDebutAttest)+'" AND PPU_DATEFIN<="'+UsDateTime(DateFinAttest)+'"'+
           ' AND (PPU_DATEDEBUT NOT IN (SELECT PCI_DEBUTCONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+LeSalarie+'"'+
           ' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDebutAttest)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFinAttest)+'" AND PCI_FINCONTRAT=PPU_DATEFIN)'+
           ' OR PPU_DATEFIN NOT IN (SELECT PCI_FINCONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+LeSalarie+'"'+
           ' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDebutAttest)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFinAttest)+'" AND PCI_DEBUTCONTRAT=PPU_DATEDEBUT))'+
           ' ORDER BY PPU_DATEDEBUT',True);
        TobTemp := Tob.Create('Tobtemp',Nil,-1);
        TobTemp.LoadDetailDB('Tobtemp','','',Q,False);
        Ferme(Q);
        For x := 0 to TobTemp.detail.Count - 1 do
        begin
                TLP := Tob.Create('UnePeriode',TobLesPeriodes,-1);
                DD := TobTemp.Detail[x].GetValue('PPU_DATEDEBUT');
                DF := TobTemp.Detail[x].GetValue('PPU_DATEFIN');
                BrutAv := TobTemp.Detail[x].GetValue('PPU_CBRUT');
                BrutAp := TobTemp.Detail[x].GetValue('PPU_CBRUT');
                Emploi := ''; //PT18
                NumObjet := ''; //PT28

                //PT29 Q := OpenSQL('SELECT PCI_LIBELLEEMPLOI FROM CONTRATTRAVAIL WHERE PCI_DEBUTCONTRAT="'+UsDateTime(DD)+'" OR PCI_FINCONTRAT="'+UsDateTime(DF)+'"',True);
                Q := OpenSQL('SELECT PCI_LIBELLEEMPLOI FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+LeSalarie+'" AND '+
                '(PCI_DEBUTCONTRAT="'+UsDateTime(DD)+'" OR PCI_FINCONTRAT="'+UsDateTime(DF)+'")',True);

                If Not Q.Eof then Emploi := Q.FindField('PCI_LIBELLEEMPLOI').AsString;
                Ferme(Q);

                //PT28 Debut Ajout ===>
                Q := OpenSQL('SELECT PCI_NUMOBJETAEM FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+LeSalarie+'" AND '+
                '(PCI_DEBUTCONTRAT <= "'+UsDateTime(DD)+'" AND PCI_FINCONTRAT >= "'+UsDateTime(DF)+'")',True);

                If Not Q.Eof then NumObjet := Q.FindField('PCI_NUMOBJETAEM').AsString;
                Ferme(Q);
                //PT28 Fin Ajout <===

                Q := OpenSQL('SELECT SUM(PHB_MTREM) MONTANT FROM HISTOBULLETIN '+
                'LEFT JOIN REMUNERATION ON PHB_RUBRIQUE=PRM_RUBRIQUE '+
                'WHERE PHB_SALARIE="'+LeSalarie+'" AND PHB_DATEDEBUT="'+UsDateTime(DD)+'" '+
                'AND PHB_DATEFIN="'+UsdateTime(DF)+'" AND PRM_THEMEREM="ABT"',True);
                If Not Q.Eof then BrutAp := BrutAv - Q.FindField('MONTANT').AsFloat;
                Ferme(Q);
                //DEBUT PT13
                Q := OpenSQL('SELECT SUM(('+ChampBul+')) MONTANT FROM HISTOBULLETIN WHERE PHB_SALARIE="'+LeSalarie+'" '+
                'AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+RubJours+'" '+
                'AND PHB_DATEDEBUT>="'+UsDateTime(DD)+'" '+
                'AND PHB_DATEFIN<="'+UsDateTime(DF)+'"',True);
                If Not Q.Eof then Jours := Q.FindField('MONTANT').AsFloat
                else Jours := 0;
                Ferme(Q);
                //FIN PT13
                TLP.AddChampSupValeur('ORIGINE','P');
                //DEBUT PT15
                DDc := Idate1900;
                DFc := Idate1900;
                Q := OpenSQL('SELECT PCI_DEBUTCONTRAT,PCI_FINCONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+LeSalarie+'" '+
                'AND PCI_DEBUTCONTRAT<="'+UsDateTime(DD)+'" AND PCI_FINCONTRAT>="'+UsdateTime(DF)+'"',True);
                If Not Q.Eof then
                begin
                        DDc := Q.FindField('PCI_DEBUTCONTRAT').AsdateTime;
                        DFc := Q.FindField('PCI_FINCONTRAT').AsdateTime;
                end;
                Ferme(Q);
                //FIN PT15
                If DDc = Idate1900 then   // PT19
                begin
                     DDc := TobTemp.Detail[x].GetValue('PSA_DATEENTREE');
                     DFc := TobTemp.Detail[x].GetValue('PSA_DATESORTIE');
                end;
                TLP.AddChampSupValeur('DATEDEBUT',DDc);
                TLP.AddChampSupValeur('DATEFIN',DFc);
                TLP.AddChampSupValeur('LIBDATEDEBUT',DD);
                TLP.AddChampSupValeur('LIBDATEFIN',DF);
                TLP.AddChampSupValeur('NBCACHETS',0);
                TLP.AddChampSupValeur('NBHEURES',0);
                TLP.AddChampSupValeur('NBJOURS',Jours);
                TLP.AddChampSupValeur('BRUT',BrutAV);
                TLP.AddChampSupValeur('NET',BrutAp);
                TLP.AddChampSupValeur('RUPTURE','FIN');//PT5
                TLP.AddChampSupValeur('EMPLOI',Emploi);  //PT18
                TLP.AddChampSupValeur('NOBJET',NumObjet);  //PT28

                Q := OpenSQL('SELECT '+ChampBulC+' FROM HISTOBULLETIN '+
                        'WHERE PHB_SALARIE="'+LeSalarie+'" AND PHB_DATEDEBUT="'+UsDateTime(DD)+'" '+
                        'AND PHB_DATEFIN="'+UsdateTime(DF)+'" AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+RubCachets+'"',True);
                If Not Q.Eof then TLP.PutValue('NBCACHETS',Q.FindField(ChampBulC).AsFloat)
                else TLP.Putvalue('NBHEURES',TobTemp.Detail[x].GetValue('PPU_CHEURESTRAV'));
                Ferme(Q);
        end;
        TobTemp.Free;
        //PT28 Ajout champ PCI_NUMOBJETAEM dans le select
        Q := OpenSQL('SELECT PCI_LIBELLEEMPLOI,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,PCI_ISHEURES,PCI_ISCACHETS,PCI_ISNBEFFECTUE,PCI_MOTIFSORTIE,PCI_NUMOBJETAEM FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+LeSalarie+'"'+
           ' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDebutAttest)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFinAttest)+'"'+
           ' AND PCI_FINCONTRAT>="'+UsDateTime(DateDebutAttest)+'" AND PCI_FINCONTRAT<="'+UsDateTime(DateFinAttest)+'"'+
           ' AND (PCI_DEBUTCONTRAT NOT IN (SELECT PPU_DATEDEBUT FROM PAIEENCOURS WHERE PPU_SALARIE="'+LeSalarie+'"'+
           ' AND PPU_DATEDEBUT>="'+UsDateTime(DateDebutAttest)+'" AND PPU_DATEDEBUT<="'+UsDateTime(DateFinAttest)+'" AND PCI_FINCONTRAT=PPU_DATEFIN) '+
           'OR PCI_FINCONTRAT NOT IN (SELECT PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_SALARIE="'+LeSalarie+'"'+
           ' AND PPU_DATEDEBUT>="'+UsDateTime(DateDebutAttest)+'" AND PPU_DATEDEBUT<="'+UsDateTime(DateFinAttest)+'" AND PCI_DEBUTCONTRAT=PPU_DATEDEBUT)) '+
           'ORDER BY PCI_DEBUTCONTRAT',True);
        TobTemp := Tob.Create('Tobtemp',Nil,-1);
        TobTemp.LoadDetailDB('Tobtemp','','',Q,False);
        Ferme(Q);
        For x := 0 to TobTemp.detail.Count - 1 do
        begin
                DD := TobTemp.Detail[x].GetValue('PCI_DEBUTCONTRAT');
                DF := TobTemp.Detail[x].GetValue('PCI_FINCONTRAT');
                Emploi := TobTemp.Detail[x].GetValue('PCI_LIBELLEEMPLOI'); //PT18
                NumObjet := TobTemp.Detail[x].GetValue('PCI_NUMOBJETAEM'); //PT28
                TLP := Tob.Create('UnePeriode',TobLesPeriodes,-1);
                TLP.AddChampSupValeur('ORIGINE','C');
                TLP.AddChampSupValeur('DATEDEBUT',DD);
                TLP.AddChampSupValeur('DATEFIN',DF);
                TLP.AddChampSupValeur('LIBDATEDEBUT',DD);
                TLP.AddChampSupValeur('LIBDATEFIN',DF);
                TLP.AddChampSupValeur('NBCACHETS',0);
                TLP.AddChampSupValeur('NBHEURES',0);
                TLP.AddChampSupValeur('NBJOURS',0);
                TLP.AddChampSupValeur('BRUT',0);
                TLP.AddChampSupValeur('NET',0);
                TLP.AddChampSupValeur('RUPTURE','FIN');
                TLP.AddChampSupValeur('EMPLOI',Emploi);  //PT18
                TLP.AddChampSupValeur('NOBJET',NumObjet);  //PT28

                If TobTemp.Detail[x].GetValue('PCI_MOTIFSORTIE') = '36' then TLP.PutValue('RUPTURE','EMP');
                If TobTemp.Detail[x].GetValue('PCI_MOTIFSORTIE') = '37' then TLP.PutValue('RUPTURE','SAL');
                If TobTemp.Detail[x].GetValue('PCI_ISHEURES') = 'X' then TLP.Putvalue('NBHEURES',TobTemp.Detail[x].GetValue('PCI_ISNBEFFECTUE'))
                else If TobTemp.Detail[x].GetValue('PCI_ISCACHETS') = 'X' then TLP.Putvalue('NBCACHETS',TobTemp.Detail[x].GetValue('PCI_ISNBEFFECTUE'));
        end;
        TobTemp.Free;
        If ToblesPeriodes.Detail.Count > 0 then NumPeriode := 0
        else NumPeriode := -1;
        MaxPeriode := TobLesPeriodes.Detail.Count - 1;
        If MenuPeriodes.Items.Count > 0 then MenuPeriodes.Items.Clear; //PT11
        For i := 0 to TobLesPeriodes.Detail.count - 1 do
        begin
                Origine := TobLesPeriodes.Detail[i].GetValue('ORIGINE');
                If Origine = 'C' then St := 'Contrat du '+dateToStr(TobLesPeriodes.Detail[i].GetValue('LIBDATEDEBUT')) +' au '+ DateToStr(TobLesPeriodes.Detail[i].GetValue('LIBDATEFIN'))
                else if Origine = 'P' then St := 'Paie du '+dateToStr(TobLesPeriodes.Detail[i].GetValue('LIBDATEDEBUT')) +' au '+ DateToStr(TobLesPeriodes.Detail[i].GetValue('LIBDATEFIN'))
                else St := 'Paie et contrat du '+dateToStr(TobLesPeriodes.Detail[i].GetValue('LIBDATEDEBUT')) +' au '+ DateToStr(TobLesPeriodes.Detail[i].GetValue('LIBDATEFIN'));
                NewItems := TMenuItem.Create(MenuPeriodes);
                NewItems.Caption := St;
                NewItems.Name := 'N'+ IntToStr(i);
                MenuPeriodes.Items.Add(NewItems);
                MenuPeriodes.Items[i].OnClick := MenuPeriodesClick;
        end;
        BPeriode := TToolBarButton97(GetControl('BTPER'));
        BPeriode.Refresh;
end;

Function TOF_PGASSEDIC_SPECTACLE.CalculModulo97(NumAttest : String) : Integer;
var Nb,L,i : Integer;
    Alpha,CleAlpha : string;
    Ch : Char;
begin
     Alpha := Copy(NumAttest,2,3);
     L := Length(Alpha);
     CleAlpha := '';
     For i := 1 to L do
     begin
          Ch := Alpha[i];
          If IsNumeric(Ch) then CleAlpha := CleAlpha + Ch
          else if Ch in ['A','J'] then CleAlpha := CleAlpha + '1'
          else if Ch in ['B','K','S'] then CleAlpha := CleAlpha + '2'
          else if Ch in ['C','L','T'] then CleAlpha := CleAlpha + '3'
          else if Ch in ['D','M','U'] then CleAlpha := CleAlpha + '4'
          else if Ch in ['E','N','V'] then CleAlpha := CleAlpha + '5'
          else if Ch in ['F','O','W'] then CleAlpha := CleAlpha + '6'
          else if Ch in ['G','P','X'] then CleAlpha := CleAlpha + '7'
          else if Ch in ['H','Q','Y'] then CleAlpha := CleAlpha + '8'
          else if Ch in ['I','R','Z'] then CleAlpha := CleAlpha + '9';
     end;

        If IsNumeric(Copy(NumAttest,5,6)) then
        begin
             Nb := StrToInt(CleAlpha+Copy(NumAttest,5,6));
             Result := Nb mod 97;
        end
        else Result := 0;
end;

Procedure TOF_PGASSEDIC_SPECTACLE.GereCheckBox (Sender : TObject);
begin
        if TCheckBox(Sender) = Nil then Exit;
        if TCheckBox(Sender).checked = False then exit;
        If TCheckBox(Sender).Name = 'CATTESTINIT' then
        begin
                If GetCheckBoxState('CATTESTINIT') = CbChecked then
                begin
                        SetControlChecked('ATTESTCOMPL',False);
                        SetControlChecked('CRECAPPOS',False);
                        SetControlChecked('CRECAPNEG',False);
                        SetControlText('NUMATTESTINIT','');
                        SetControlEnabled('NUMATTESTINIT',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'ATTESTCOMPL' then
        begin
                If GetCheckBoxState('ATTESTCOMPL') = CbChecked then
                begin
                        SetControlChecked('CATTESTINIT',False);
                        SetControlChecked('CRECAPPOS',False);
                        SetControlChecked('CRECAPNEG',False);
                        SetControlEnabled('NUMATTESTINIT',True);
                        SetFocusControl('NUMATTESTINIT');   //PT26
                end;
        end;
        If TCheckBox(Sender).Name = 'CRECAPPOS' then
        begin
                If GetCheckBoxState('CRECAPPOS') = CbChecked then
                begin
                        SetControlChecked('ATTESTCOMPL',False);
                        SetControlChecked('CATTESTINIT',False);
                        SetControlChecked('CRECAPNEG',False);
                        SetControlEnabled('NUMATTESTINIT',True);
                        SetFocusControl('NUMATTESTINIT');   //PT26
                end;
        end;
        If TCheckBox(Sender).Name = 'CRECAPNEG' then
        begin
                If GetCheckBoxState('CRECAPNEG') = CbChecked then
                begin
                        SetControlChecked('ATTESTCOMPL',False);
                        SetControlChecked('CRECAPPOS',False);
                        SetControlChecked('CATTESTINIT',False);
                        SetControlEnabled('NUMATTESTINIT',True);
                        SetFocusControl('NUMATTESTINIT');   //PT26
                end;
        end;
        If TCheckBox(Sender).Name = 'CSALCADRE' then
        begin
                If GetCheckBoxState('CSALCADRE') = CbChecked then SetControlChecked('CSALNONCADRE',False);
        end;
        If TCheckBox(Sender).Name = 'CSALNONCADRE' then
        begin
                If GetCheckBoxState('CSALNONCADRE') = CbChecked then SetControlChecked('CSALCADRE',False);
        end;
        If TCheckBox(Sender).Name = 'CARTISTE' then
        begin
                If GetCheckBoxState('CARTISTE') = CbChecked then
                begin
                        SetControlChecked('CREALISATEUR',False);  //PT26
                        SetControlChecked('CTECHNICIEN',False);
                        SetControlChecked('COUVRIER',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'COUVRIER' then
        begin
                If GetCheckBoxState('COUVRIER') = CbChecked then
                begin
                        SetControlChecked('CREALISATEUR',False);  //PT26
                        SetControlChecked('CTECHNICIEN',False);
                        SetControlChecked('CARTISTE',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'CTECHNICIEN' then
        begin
                If GetCheckBoxState('CTECHNICIEN') = CbChecked then
                begin
                        SetControlChecked('CREALISATEUR',False);  //PT26
                        SetControlChecked('CARTISTE',False);
                        SetControlChecked('COUVRIER',False);
                end;
        end;
        //PT26 =============>
        If TCheckBox(Sender).Name = 'CREALISATEUR' then
        begin
                If GetCheckBoxState('CREALISATEUR') = CbChecked then
                begin
                        SetControlChecked('CTECHNICIEN',False);
                        SetControlChecked('CARTISTE',False);
                        SetControlChecked('COUVRIER',False);
                end;
        end;
        //PT26 <=============
        if (GetCheckBoxState('CTECHNICIEN')=CbChecked) or (GetCheckBoxState('COUVRIER')=CbChecked) then
        begin
           SetControlEnabled('HEURESART',False);
           SetControlText('HEURESART','0');
           //PT26 SetControlEnabled('HEURESOUV',True);
           SetControlEnabled('CACHETS',False);   //PT26
           SetControlText('CACHETS','0');        //PT26
        end
        else
        begin
           //PT26 SetControlEnabled('HEURESOUV',False);
           SetControlEnabled('HEURESART',True);
           SetControlEnabled('CACHETS',True);   //PT26
           //PT26 SetControlText('HEURESOUV','0');
        end;
        If TCheckBox(Sender).Name = 'CLICENCE' then
        begin
                If GetCheckBoxState('CLICENCE') = CbChecked then
                begin
                        SetControlChecked('CLICENCENON',False);
                        SetControlEnabled('ENUMLICENCE',True);
                        SetFocusControl('ENUMLICENCE');   //PT26
                end;
        end;
        If TCheckBox(Sender).Name = 'CLICENCENON' then
        begin
                If GetCheckBoxState('CLICENCENON') = CbChecked then
                begin
                        SetControlChecked('CLICENCE',False);
                        SetControlEnabled('ENUMLICENCE',False);
                        SetControlText('ENUMLICENCE','');
                end;
        end;
        If TCheckBox(Sender).Name = 'CORGANISATEUR' then
        begin
                If GetCheckBoxState('CORGANISATEUR') = CbChecked then SetControlChecked('CORGANISATEURNON',False);
        end;
        If TCheckBox(Sender).Name = 'CORGANISATEURNON' then
        begin
                If GetCheckBoxState('CORGANISATEURNON') = CbChecked then SetControlChecked('CORGANISATEUR',False);
        end;
        If TCheckBox(Sender).Name = 'CLABEL' then
        begin
                If GetCheckBoxState('CLABEL') = CbChecked then
                begin
                        SetControlChecked('CLABELNON',False);
                        SetControlEnabled('ENUMLABEL',True);
                        SetFocusControl('ENUMLABEL');   //PT26
                end;
        end;
        If TCheckBox(Sender).Name = 'CLABELNON' then
        begin
                If GetCheckBoxState('CLABELNON') = CbChecked then
                begin
                        SetControlChecked('CLABEL',False);
                        SetControlEnabled('ENUMLABEL',False);
                        SetControlText('ENUMLABEL','');
                end;
        end;
{       PT22 Mis en commentaire car inaccessible
        If TCheckBox(Sender).Name = 'CAFFILCP' then
        begin
                If GetCheckBoxState('CAFFILCP') = CbChecked then
                begin
                        SetControlChecked('CAFFILCPNON',False);
                        SetControlEnabled('ENUMAFFILIATION',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'CAFFILCPNON' then
        begin
                If GetCheckBoxState('CAFFILCPNON') = CbChecked then
                begin
                        SetControlChecked('CAFFILCP',False);
                        SetControlEnabled('ENUMAFFILIATION',False);
                        SetControlText('ENUMAFFILIATION','');
                end;
        end;       }
        If TCheckBox(Sender).Name = 'PARENTE' then
        begin
                If GetCheckBoxState('PARENTE') = CbChecked then
                begin
                        SetControlChecked('PARENTENON',False);
                        SetControlEnabled('LIENPARENTE',True);
                        SetFocusControl('LIENPARENTE');   //PT26
                end;
        end;
        If TCheckBox(Sender).Name = 'PARENTENON' then
        begin
                If GetCheckBoxState('PARENTENON') = CbChecked then
                begin
                        SetControlChecked('PARENTE',False);
                        SetControlText('LIENPARENTE','');
                        SetControlEnabled('LIENPARENTE',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'CFINCDD' then
        begin
                If GetCheckBoxState('CFINCDD') = CbChecked then
                begin
                        SetControlChecked('CRUPTEMPL',False);
                        SetControlChecked('CRUPTSAL',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'CRUPTEMPL' then
        begin
                If GetCheckBoxState('CRUPTEMPL') = CbChecked then
                begin
                        SetControlChecked('CRUPTSAL',False);
                        SetControlChecked('CFINCDD',False);
                end;
        end;
        If TCheckBox(Sender).Name = 'CRUPTSAL' then
        begin
                If GetCheckBoxState('CRUPTSAL') = CbChecked then
                begin
                        SetControlChecked('CRUPTEMPL',False);
                        SetControlChecked('CFINCDD',False);
                end;
        end;
        //PT26 ========>
        If TCheckBox(Sender).Name = 'CENCOURS' then
        begin
                If GetCheckBoxState('CENCOURS') = CbChecked then
                begin
                        SetControlChecked('CRUPTEMPL',False);
                        SetControlChecked('CFINCDD',False);
                        SetControlChecked('CRUPTSAL',False);
                        SetControlText('EDATEFIN',DateToStr(IDate1900));
                end;
        End;
        //PT26 <========
end;

Procedure TOF_PGASSEDIC_SPECTACLE.AfficherSalarie;
var Q:TQuery;
    SEmploi : String;
    QAttes : TQuery;
    CategIS : String;
begin
SetControlCaption('LMATRICULE',LeSalarie);
SEmploi:='';
CategIS:='';
SetControlChecked('PARENTENON',True);
Q:=OpenSQL('SELECT PSA_NOMJF,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,PSA_LIBELLEEMPLOI,PSA_DADSCAT,PSE_ISRETRAITE,PSE_ISCATEG,PSE_ISNUMASSEDIC,PSE_ISNUMIDENT,'+
          'PSA_DATENAISSANCE,PSA_TELEPHONE,PSA_CODEPOSTAL,PSA_VILLE,PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,PSA_SURNOM,PSA_DATENAISSANCE,PSA_NUMEROSS'+
        ' FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE WHERE PSA_SALARIE="'+LeSalarie+'"',True);
If Not Q.eof Then                             //PortageCWAS
   begin
         //DEBUT PT24
         If Q.FindField('PSA_NOMJF').AsString = '' then
         begin
            SetControlText('LNOM',Q.FindField('PSA_LIBELLE').AsString);
            SetControlText('ENOMUSAGE',Q.FindField('PSA_SURNOM').AsString);
         end
         else
         begin
            SetControlText('LNOM',Q.FindField('PSA_NOMJF').AsString);
            SetControlText('ENOMUSAGE',Q.FindField('PSA_LIBELLE').AsString);
         end;
        //FIN PT24
         SetControlText('LPRENOM',Q.FindField('PSA_PRENOM').AsString);
         SEmploi:=Q.FindField('PSA_LIBELLEEMPLOI').AsString;
         EmploiSalarie := Q.FindField('PSA_LIBELLEEMPLOI').AsString; //PT18
         CategIS:=Q.FindField('PSE_ISCATEG').AsString;
         LaCategorie := Q.FindField('PSE_ISCATEG').AsString; //PT17
         SetControlText('ERETRAITE',Q.FindField('PSE_ISRETRAITE').AsString);
         //PT28 SetControlText('ENUMASSEDIC',Q.FindField('PSE_ISNUMASSEDIC').AsString);
         SetControlText('ENUMIDENTIFIANT',Q.FindField('PSA_NUMEROSS').AsString); //PT8
         SetControlEnabled('ENUMIDENTIFIANT',False); //PT8
         SetControlText('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));
{PT27
         SetControlText('ETELSALARIE',Q.FindField('PSA_TELEPHONE').AsString);
}
         SetControlText('ECPSAL',Q.FindField('PSA_CODEPOSTAL').AsString);
         SetControlText('ECOMMUNESAL',Q.FindField('PSA_VILLE').AsString);
         SetControlText('EADRSAL1',Q.FindField('PSA_ADRESSE1').AsString);
         SetControlText('EADRSAL2',Q.FindField('PSA_ADRESSE2').AsString + Q.FindField('PSA_ADRESSE3').AsString);
         //PT26 Dejà Fait SetControlText('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));
         EtabSalarie := Q.FindField('PSA_ETABLISSEMENT').AsString;
         If Q.FindField('PSA_DADSCAT').AsString='01' then SetControlChecked('CSALCADRE',True) //PT14
         else SetControlChecked('CSALNONCADRE',True);
   end;
Ferme(Q);
If CategIS='OUV' Then SetControlChecked('COUVRIER',True)
else If (CategIS='TEC') Then SetControlChecked('CTECHNICIEN',True)
else If CategIS='ART' Then SetControlChecked('CARTISTE',True)
else If CategIS='REA' Then SetControlChecked('CREALISATEUR',True);  //PT26

SetControlText('EEMPLOI',RechDom('PGLIBEMPLOI',SEmploi,False));
//DEBUT PT11
{DecodeDate(DateDecl,yy,mm,dd);
Mois := IntToStr(mm);
Annee := IntToStr(yy);
SetControlText('MOIS',Mois);
SetControlText('ANNEE',Annee);}
//FIN PT11
QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + EtabSalarie + '%") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ASP%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
if not QAttes.eof then
begin
        SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
        AfficheDeclarant(nil);
end;
Ferme(QAttes);
SetControlText('ETAUX',FloatToStr(TauxEltNat));  // PT7
ChargeZonesEmployeur;
If Action = 'CREATION' then
begin
        RemplirTobLesPeriodes;
        AfficherPeriode(NumPeriode);
end
else RecupAttestation;
SetFocusControl('MOIS'); //PT26  pour se positionner en début
end;

Procedure TOF_PGASSEDIC_SPECTACLE.AfficherPeriode(NumFille : Integer);
var T : Tob;
    DD : TDateTime;
    aa,mm,jj : Word;
    Mois : String;
begin
        SetControlText('ETAUX',FloatToStr(TauxEltNat));
        SetControlchecked('CATTESTINIT',True);
        If NumFille = -1 then
        begin
                DD := Date;
                DecodeDate(DD,aa,mm,jj);
                SetControlText('ANNEE',CodeAnnee(aa));
                Mois := IntToStr(mm);
                If Length(Mois) = 1 then Mois := '0'+Mois;
                SetControlText('MOIS',Mois);
                SetControlText('EDATEDEB',DateToStr(Date));
                SetControlText('EDATEFIN',DateToStr(Date));
                PGIBox('Aucune période n''a été trouvée pour ce salarié',Ecran.Caption);
        //DEBUT PT6
                SetControlEnabled('BPERSUIV',False);
                SetControlEnabled('BPERPREC',False);
                SetControlEnabled('BTPER',False);
                Exit;
        end
        else
        begin
                SetControlEnabled('BPERSUIV',True);
                SetControlEnabled('BPERPREC',True);
                SetControlEnabled('BTPER',True);
        end;
        //FIN PT6
        If NumFille = MaxPeriode then SetControlEnabled('BPERSUIV',False)
        else SetControlEnabled('BPERSUIV',True);
        If NumFille = 0 then SetControlEnabled('BPERPREC',False)
        else SetControlEnabled('BPERPREC',True);
        T := TobLesPeriodes.Detail[NumFille];
        If T <> Nil then
        begin
                DD := T.GetValue('LIBDATEDEBUT');
                DecodeDate(DD,aa,mm,jj);
                SetControlText('ANNEE',CodeAnnee(aa));
                Mois := IntToStr(mm);
                If Length(Mois) = 1 then Mois := '0'+Mois;
                SetControlText('MOIS',Mois);
                SetControlText('EDATEDEB',DateToStr(T.GetValue('DATEDEBUT')));
                SetControlText('EDATEFIN',DateToStr(T.GetValue('DATEFIN')));
                SetControltext('ESALBTRUTSAV',FloatToStr(T.GetValue('BRUT')));
                SetControltext('ESALBTRUTSAP',FloatToStr(T.GetValue('NET')));
                SetControltext('NBJOURSTRAV',FloatToStr(T.GetValue('NBJOURS')));
                If (LaCategorie = 'REA') or (GetCheckBoxState('CARTISTE') = CbChecked) then //PT17
                begin
                        //PT26 Debut modif =====>
                        //SetControlText('CACHETS',T.GetValue('NBCACHETS'));
                        //SetControlText('HEURESART',T.GetValue('NBHEURES'));
                        //SetControlText('HEURESOUV','0');
                        SetControlText('HEURESOUV',T.GetValue('NBHEURES'));
                        If StrToFloat(GetControltext('NBJOURSTRAV')) < 5 Then
                        Begin
                           SetControltext('HEURESART',T.GetValue('NBCACHETS'));
                           SetControltext('CACHETS','0');
                        End
                        Else
                        Begin
                           SetControltext('CACHETS',T.GetValue('NBCACHETS'));
                           SetControltext('HEURESART','0');
                        End;
                        //PT26 Fin modif <=====
                end
                else If (GetCheckBoxState('CTECHNICIEN') = CbChecked) or (GetCheckBoxState('COUVRIER') = CbChecked) then
                begin
                        //PT20 - Début
                        {
                        If (GetCheckBoxState('CTECHNICIEN') = CbChecked) then
                        begin
                             SetControlText('CACHETS',T.GetValue('NBCACHETS'));
                             SetControlText('HEURESART',T.GetValue('NBHEURES'));
                             SetControlText('HEURESOUV','0');
                        end
                        else
                        begin}
                             SetControlText('HEURESOUV',T.GetValue('NBHEURES'));
                             SetControlText('CACHETS','0');
                             SetControlText('HEURESART','0');
                        {end;}
                        //PT20 - Fin
                end
                else
                begin
                        SetControlText('HEURESOUV',T.GetValue('NBHEURES'));
                        //PT26 Debut modif =====>
                        //SetControlText('CACHETS',T.GetValue('NBCACHETS'));
                        //SetControlText('HEURESART',T.GetValue('NBHEURES'));
                        If StrToFloat(GetControltext('NBJOURSTRAV')) < 5 Then
                        Begin
                           SetControltext('HEURESART',T.GetValue('NBCACHETS'));
                           SetControltext('CACHETS','0');
                        End
                        Else
                        Begin
                           SetControltext('CACHETS',T.GetValue('NBCACHETS'));
                           SetControltext('HEURESART','0');
                        End;
                        //PT26 Fin modif <=====
                end;
                SetControlChecked('CFINCDD',(T.GetValue('RUPTURE') = 'FIN'));
                SetControlChecked('CRUPTEMPL',(T.GetValue('RUPTURE') = 'EMP'));
                SetControlChecked('CRUPTSAL',(T.GetValue('RUPTURE') = 'SAL'));
                If T.GetValue('EMPLOI') <> '' then SetControlText('EEMPLOI',RechDom('PGLIBEMPLOI',T.GetValue('EMPLOI'),False))
                else SetControlText('EEMPLOI',RechDom('PGLIBEMPLOI',EmploiSalarie,False)); //PT18
                SetControlText('NUMOBJET',T.GetValue('NOBJET'));  //PT28
        end;
        ChangeMoisAnnee(Nil);
end;

Procedure TOF_PGASSEDIC_SPECTACLE.PeriodeSuivante(Sender : TObject);
begin
MetABlanc(False,True);
NumPeriode := NumPeriode + 1;
AfficherPeriode(NumPeriode);
end;

Procedure TOF_PGASSEDIC_SPECTACLE.PeriodePrecedente(Sender : TObject);
begin
MetABlanc(False,True);
NumPeriode := NumPeriode - 1;
AfficherPeriode(NumPeriode);
end;

procedure TOF_PGASSEDIC_SPECTACLE.AfficheDeclarant(Sender : Tobject);
var
  St: string;
begin
  if GetControlText('DECLARANT') = '' then exit;
  St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
  SetControlText('ENOMEMPL', RechDom('PGDECLARANTNOM', GetControlText('DECLARANT'), False));
  SetControlText('EPRENOMEMPL', RechDom('PGDECLARANTPRENOM', GetControlText('DECLARANT'), False));
  SetControlText('ELIEU', RechDom('PGDECLARANTVILLE', GetControlText('DECLARANT'), False));
  SetControlText('PERSAJOINDRE', RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False));
  SetControlText('TELPERSAJOINDRE', RechDom('PGDECLARANTTEL', GetControlText('DECLARANT'), False));
  St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
  if St = 'AUT' then SetControlText('EQUALITEEMPL', RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False))
  else SetControlText('EQUALITEEMPL', RechDom('PGQUALDECLARANT2', St, False));
end;

Procedure TOF_PGASSEDIC_SPECTACLE.MenuPeriodesClick(Sender : TObject);
var Libelle,IndexPeriode : String;
    NbChar : Integer;
begin
        Libelle := TMenuItem(Sender).Name;
        NbChar := Length(Libelle);
        IndexPeriode := Copy(Libelle,2,NbChar-1);
        MetABlanc(False,True);
        NumPeriode := StrToInt(IndexPeriode);
        AfficherPeriode(NumPeriode);
end;

procedure TOF_PGASSEDIC_SPECTACLE.ChangeMoisAnnee (Sender : TObject);
var mm,yy : word;
    DatePeriodeEC : TDateTime;
    An : String;
begin
        If getControlText('MOIS') = '' then Exit;
        MM := StrToInt(GetcontrolText('MOIS'));
        An := RechDom('PGANNEE',GetControlText('ANNEE'),False);
        If (IsValidDate(GetControlText('EDATEFIN'))) and (Isnumeric(An)) and (length(An) = 4) then
        begin
                yy := StrToInt(An);
                DatePeriodeEC := EncodeDate(yy,mm,01);
                DatePeriodeEC := FinDeMois(DatePeriodeEC);
                If StrToDate(GetControlText('EDATEFIN')) > DatePeriodeEC then
                begin
                     SetControlChecked('CENCOURS',True);
                     //DEBUT PT21
                     SetControlChecked('CFINCDD',False);
                     SetControlChecked('CRUPTEMPL',False);
                     SetControlChecked('CRUPTSAL',False);
                     //FIN PT21
                end
                else SetControlChecked('CENCOURS',False);
        end;
        RecupNumAttestInit;
        //If Action = 'CREATION' then RecupDerNumObjet; //PT26 PT28
end;

procedure TOF_PGASSEDIC_SPECTACLE.AccesNumAttest(Sender : TObject);
begin
     SetControlEnabled('ENUMDOC',TCheckBox(Sender).Checked = False);
     If GetCheckBoxState('CINCREMENTAUTO') = CbUnChecked Then SetFocusControl('ENUMDOC');   //PT26
end;

procedure TOF_PGASSEDIC_SPECTACLE.ValidationAttest (Sender : Tobject);
//PT26 var NumParam,NumAttest,Agrement,Increment : String;
var NumParam,Agrement,Increment : String;
    TobAEM,TA : Tob;
    Q : TQuery;
    Modulo : Integer;
    AjoutV4,C_Modulo : String; //PT26
begin
  //PT28 If CtrlSaisie = False Then Exit;  //PT26
  If CtrlSaisie(TRUE) = False Then Exit;   //PT28

  TobAEM := Tob.Create('PGAEM',Nil,-1);
  If Action = 'CREATION' then
  begin
       If GetCheckBoxState('CINCREMENTAUTO') = CbChecked then
       begin
         If Length(GetParamSoc('SO_PGINCREMENTAEM')) = 9 then //PT10
         begin
               NumParam := GetParamSoc('SO_PGINCREMENTAEM');
               Agrement := Copy(NumParam,1,3);
               Increment := Copy(NumParam,4,6);
               Increment := IntToStr(StrToInt(Increment)+ 1);
               Increment := ColleZeroDevant(StrToInt(Increment),6);
               Modulo := CalculModulo97('B'+Agrement+Increment);
               If Modulo < 10 Then C_Modulo := '0'+ IntToStr(Modulo) Else C_Modulo := IntToStr(Modulo); //PT26
               //SetControlText('ENUMDOC','B'+Agrement+Increment+IntToStr(Modulo));                       PT26
               SetControlText('ENUMDOC','B'+Agrement+Increment+C_Modulo);                               //PT26
         end;
         NumParam := COPY(GetControlText('ENUMDOC'),2,9);
         SetParamSoc('SO_PGINCREMENTAEM',NumParam);  //PT10
         {$IFDEF EAGLCLIENT}
         AvertirCacheServer('PARAMSOC');
         {$ENDIF}
       end;
       TA := Tob.Create('PGAEM',TobAEM,-1);
       TA.PutValue('PGA_NUMATTEST',GetControlText('ENUMDOC'));
  end
  else
  begin
        Q := OpenSQL('SELECT * FROM PGAEM WHERE PGA_NUMATTEST="'+GetControlText('ENUMDOC')+'"',True);
        TobAEM.LoadDetailDB('PGAEM','','',Q,False);
        Ferme(Q);
        TA := TobAEM.FindFirst(['PGA_NUMATTEST'],[getControlText('ENUMDOC')],False);
  end;
  TA.PutValue('PGA_SALARIE',LeSalarie);
  TA.PutValue('PGA_ETABLISSEMENT',EtabSalarie);
  If GetCheckBoxState('CATTESTINIT') = CbChecked then TA.PutValue('PGA_TYPEAEM','INI')
  else If GetCheckBoxState('ATTESTCOMPL') = CbChecked then TA.PutValue('PGA_TYPEAEM','COM')
  else If GetCheckBoxState('CRECAPPOS') = CbChecked then TA.PutValue('PGA_TYPEAEM','POS')
  else If GetCheckBoxState('CRECAPNEG') = CbChecked then TA.PutValue('PGA_TYPEAEM','NEG');
  TA.PutValue('PGA_DATEATTEST',StrToDate(GetControlText('EDATE')));
  TA.PutValue('PGA_LIBELLEEMPLOI',GetControlText('EEMPLOI'));
  TA.PutValue('PGA_NBHEURES',StrToFloat(GetControlText('HEURESOUV'))); //PT26

  If GetCheckBoxState('CARTISTE') = CbChecked then
  begin
        TA.PutValue('PGA_ISCATEG','ART');
        //PT26 TA.PutValue('PGA_NBHEURES',StrToFloat(GetControlText('HEURESART')));
        //PT26 Debut Ajout ====>
        If StrToFloat(GetControlText('NBJOURSTRAV')) < 5 Then
           TA.PutValue('PGA_NBCACHETS',StrToFloat(GetControlText('HEURESART')))
           Else TA.PutValue('PGA_NBCACHETS',StrToFloat(GetControlText('CACHETS')));
        //PT26 Fin Ajout <====
  end
  else If GetCheckBoxState('COUVRIER') = CbChecked then
  begin
        TA.PutValue('PGA_ISCATEG','OUV');
        //PT26 TA.PutValue('PGA_NBHEURES',StrToFloat(GetControlText('HEURESOUV')));
        TA.PutValue('PGA_NBCACHETS',0); //PT26
  end
  else If GetCheckBoxState('CTECHNICIEN') = CbChecked then
  begin
        TA.PutValue('PGA_ISCATEG','TEC');
        //PT26 TA.PutValue('PGA_NBHEURES',StrToFloat(GetControlText('HEURESOUV')))
        TA.PutValue('PGA_NBCACHETS',0); //PT26
  end
  Else
  begin
        //PT26 TA.PutValue('PGA_NBHEURES',StrToFloat(GetControlText('HEURESART')));
        TA.PutValue('PGA_ISCATEG','REA');
        //PT26 Debut Ajout ====>
        If StrToFloat(GetControlText('NBJOURSTRAV')) < 5 Then
           TA.PutValue('PGA_NBCACHETS',StrToFloat(GetControlText('HEURESART')))
           Else TA.PutValue('PGA_NBCACHETS',StrToFloat(GetControlText('CACHETS')));
        //PT26 Fin Ajout <====
  end;
  TA.PutValue('PGA_ISRETRAITE',GetControlText('ERETRAITE'));
  If GetCheckBoxState('CSALCADRE') = CbChecked then TA.PutValue('PGA_CADRE','X')
  else TA.PutValue('PGA_CADRE','-');
  TA.PutValue('PGA_MOIS',GetControlText('MOIS'));
  TA.PutValue('PGA_ANNEE',GetControlText('ANNEE'));
  TA.PutValue('PGA_DATEDEBUT',StrToDate(GetControlText('EDATEDEB')));
  If IsValidDate(GetControlText('EDATEFIN')) then TA.PutValue('PGA_DATEFIN',StrToDate(GetControlText('EDATEFIN')))
  else TA.PutValue('PGA_DATEFIN',IDate1900);
  If GetCheckBoxState('CENCOURS') = CbChecked then TA.PutValue('PGA_CONTRATEC','X')
  else TA.PutValue('PGA_CONTRATEC','-');
  //PT26 TA.PutValue('PGA_NBCACHETS',StrToFloat(GetControlText('CACHETS')));
  TA.PutValue('PGA_NBJOURS',StrToFloat(GetControlText('NBJOURSTRAV')));
  TA.PutValue('PGA_SALAIRESAVANT',StrToFloat(GetControlText('ESALBTRUTSAV')));
  TA.PutValue('PGA_SALAIRESAPRES',StrToFloat(GetControlText('ESALBTRUTSAP')));
  TA.PutValue('PGA_TAUX',StrToFloat(GetControlText('ETAUX')));
  TA.PutValue('PGA_CONTRIBUTIONS',StrToFloat(GetControlText('ECOTISATION')));
  If GetCheckBoxState('CFINCDD') = CbChecked then TA.PutValue('PGA_FINCDD','X')
  else TA.PutValue('PGA_FINCDD','-');
  If GetCheckBoxState('CRUPTEMPL') = CbChecked then TA.PutValue('PGA_RUPTEMP','X')
  else TA.PutValue('PGA_RUPTEMP','-');
  If GetCheckBoxState('CRUPTSAL') = CbChecked then TA.PutValue('PGA_RUPTSAL','X')
  else TA.PutValue('PGA_RUPTSAL','-');
  TA.PutValue('PGA_FAIT','-');
  TA.PutValue('PGA_NUMATTESTINIT',GetControlText('NUMATTESTINIT'));
  TA.PutValue('PGA_ACTIVITEAEM','');
  //PT26 Debut Modif ==>
  //TA.PutValue('PGA_LIENPARENTE',GetControlText('LIENPARENTE'));
  AjoutV4 := GetControlText('LIENPARENTE')+ ';' + GetControlText('ESALBTRUTSAP1')+ ';' +
             GetControlText('ETAUX1')+ ';' +GetControlText('ECOTISATION1');
  TA.PutValue('PGA_LIENPARENTE',AjoutV4);
  TA.PutValue('PGA_NUMOBJETAEM',GetControlText('NUMOBJET'));
  //PT26 Fin Modif <==
  TA.InsertOrUpdateDB(False);
  TobAEM.Free;
  TFVierge(Ecran).close;  //PT26
end;

procedure TOF_PGASSEDIC_SPECTACLE.RecupAttestation;
var TobAEM : Tob;
    Q : TQuery;
    AjoutV4,Lparent,BrutSap1,Taux1,Cotisation1 : String; //PT26
begin
 Q := OpenSQL('SELECT * FROM PGAEM WHERE PGA_SALARIE="'+LeSalarie+'" AND PGA_NUMATTEST="'+LeNumAEM+'"',True);
 TobAEM := Tob.Create('UneAttest',Nil,-1);
 TobAEM.LoadDetailDB('UneAttest','','',Q,False);
 Ferme(Q);
 If TobAEM.Detail[0] <> Nil then
 begin
        If TobAEM.Detail[0].GetValue('PGA_TYPEAEM') ='INI' then SetControlChecked('CATTESTINIT',True)
        else SetControlChecked('CATTESTINIT',False);
        If TobAEM.Detail[0].GetValue('PGA_TYPEAEM') = 'COM' then SetControlChecked('ATTESTCOMPL',True)
        else SetControlChecked('ATTESTCOMPL',False);
        if TobAEM.Detail[0].GetValue('PGA_TYPEAEM') = 'POS' then SetControlChecked('CRECAPPOS',True)
        else SetControlChecked('CRECAPPOS',False);
        if TobAEM.Detail[0].GetValue('PGA_TYPEAEM') = 'NEG' then SetControlChecked('CRECAPNEG',True)
        else SetControlChecked('CRECAPNEG',False);
        SetControlText('EDATE',DateToStr(TobAEM.Detail[0].GetValue('PGA_DATEATTEST')));
        SetControlText('EEMPLOI',TobAEM.Detail[0].GetValue('PGA_LIBELLEEMPLOI'));
        //PT26 SetControlText('HEURESART','0');
        //PT26 SetControlText('HEURESOUV','0');
        SetControlText('HEURESOUV',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBHEURES')));  //PT26
        SetControlText('NBJOURSTRAV',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBJOURS'))); //PT26

        If TobAEM.Detail[0].GetValue('PGA_ISCATEG') ='ART' then
        begin
                SetControlChecked('CARTISTE',True);
                //PT26 SetControlText('HEURESART',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBHEURES')));
                //PT26 Debut Ajout ====>
                If StrToFloat(GetControlText('NBJOURSTRAV')) < 5 Then
                Begin
                   SetControlText('HEURESART',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBCACHETS')));
                   SetControlText('CACHETS','0');
                End
                Else
                Begin
                   SetControlText('HEURESART','0');
                   SetControlText('CACHETS',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBCACHETS')));
                End;
                //PT26 Fin Ajout <====
        end
        else SetControlChecked('CARTISTE',False);

        If TobAEM.Detail[0].GetValue('PGA_ISCATEG') = 'OUV' then
        begin
                SetControlChecked('COUVRIER',True);
                //PT26 SetControlText('HEURESOUV',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBHEURES')));
                SetControlText('HEURESART','0');  //PT26
                SetControlText('CACHETS','0');    //PT26
        end
        else SetControlChecked('COUVRIER',False);

        If TobAEM.Detail[0].GetValue('PGA_ISCATEG') = 'TEC' then
        begin
                SetControlChecked('CTECHNICIEN',True);
                //PT26 SetControlText('HEURESOUV',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBHEURES')));
                SetControlText('HEURESART','0');  //PT26
                SetControlText('CACHETS','0');    //PT26
        end
        else SetControlChecked('CTECHNICIEN',False);
        //PT26 begin
             //PT26 SetControlText('HEURESART',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBHEURES')));
        //PT26 end;
        //PT26 Debut Ajout ====>
        If TobAEM.Detail[0].GetValue('PGA_ISCATEG') ='REA' then
        begin
                SetControlChecked('CREALISATEUR',True);
                If StrToFloat(GetControlText('NBJOURSTRAV')) < 5 Then
                Begin
                   SetControlText('HEURESART',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBCACHETS')));
                   SetControlText('CACHETS','0');
                End
                Else
                Begin
                   SetControlText('HEURESART','0');
                   SetControlText('CACHETS',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBCACHETS')));
                End;
        end
        else SetControlChecked('CREALISATEUR',False);
        //PT26 Fin Ajout <====

        SetControlText('ERETRAITE',TobAEM.Detail[0].GetValue('PGA_ISRETRAITE'));
        setControlChecked('CSALCADRE',TobAEM.Detail[0].GetValue('PGA_CADRE') = 'X');
        SetControlText('MOIS',TobAEM.Detail[0].GetValue('PGA_MOIS'));
        SetControlText('ANNEE',TobAEM.Detail[0].GetValue('PGA_ANNEE'));
        SetControlText('EDATEDEB',DateToStr(TobAEM.Detail[0].GetValue('PGA_DATEDEBUT')));
        SetControlText('EDATEFIN',DateToStr(TobAEM.Detail[0].GetValue('PGA_DATEFIN')));
        SetControlChecked('CENCOURS',TobAEM.Detail[0].GetValue('PGA_CONTRATEC') ='X');
        if TobAEM.Detail[0].GetValue('PGA_CONTRATEC') ='X' then RecupNumAttestInit;
        //PT26 SetControlText('CACHETS',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBCACHETS')));
        //PT26 SetControlText('NBJOURSTRAV',FloatToStr(TobAEM.Detail[0].GetValue('PGA_NBJOURS')));
        SetControlText('ESALBTRUTSAV',FloatToStr(TobAEM.Detail[0].GetValue('PGA_SALAIRESAVANT')));
        SetControlText('ESALBTRUTSAP',FloatToStr(TobAEM.Detail[0].GetValue('PGA_SALAIRESAPRES')));
        SetControlText('ETAUX',FloatToStr(TobAEM.Detail[0].GetValue('PGA_TAUX')));
        SetControlText('ECOTISATION',FloatToStr(TobAEM.Detail[0].GetValue('PGA_CONTRIBUTIONS')));
        SetControlChecked('CFINCDD',TobAEM.Detail[0].GetValue('PGA_FINCDD') = 'X');
        SetControlChecked('CRUPTEMPL',TobAEM.Detail[0].GetValue('PGA_RUPTEMP') = 'X');
        SetControlChecked('CRUPTSAL',TobAEM.Detail[0].GetValue('PGA_RUPTSAL') = 'X');
        SetControlText('NUMATTESTINIT',TobAEM.Detail[0].GetValue('PGA_NUMATTESTINIT'));
        //PT26 Debut Ajout ====>
        //SetControlText('LIENPARENTE',TobAEM.Detail[0].GetValue('PGA_LIENPARENTE'));
        //SetControlChecked('PARENTE',TobAEM.Detail[0].GetValue('PGA_LIENPARENTE') <> '');
        //SetControlChecked('PARENTENON',TobAEM.Detail[0].GetValue('PGA_LIENPARENTE') = '');
        AjoutV4 := TobAEM.Detail[0].GetValue('PGA_LIENPARENTE');
        Lparent := Trim(ReadTokenPipe(AjoutV4,';'));
        BrutSap1:= Trim(ReadTokenPipe(AjoutV4,';'));
        Taux1   := Trim(ReadTokenPipe(AjoutV4,';'));
        Cotisation1:= Trim(ReadTokenPipe(AjoutV4,';'));

        SetControlText('LIENPARENTE',Lparent);
        SetControlChecked('PARENTE',Lparent <> '');
        SetControlChecked('PARENTENON',Lparent = '');
        SetControlText('ESALBTRUTSAP1',BrutSap1);
        SetControlText('ETAUX1',Taux1);
        SetControlText('ECOTISATION1',Cotisation1);
        SetControlText('NUMOBJET',TobAEM.Detail[0].GetValue('PGA_NUMOBJETAEM'));
        //PT26 Fin Ajout <====
 end;
 TobAEM.Free;
end;

procedure TOF_PGASSEDIC_SPECTACLE.SuppAttest(Sender : TObject);
begin
        ExecuteSQL('DELETE FROM PGAEM WHERE PGA_SALARIE="'+LeSalarie+'" AND PGA_NUMATTEST="'+LeNumAEM+'"');
        TFVierge(Ecran).close;
end;
Function TOF_PGASSEDIC_SPECTACLE.CodeAnnee(An : Integer) : String;
var Q : TQuery;
begin
        Q := OpenSQL('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="PGA" AND CO_LIBELLE="'+IntToStr(An)+'"',True);
        If Not Q.Eof then Result := Q.FindField('CO_CODE').AsString
        Else Result := '';
        Ferme(Q);
end;

Procedure TOF_PGASSEDIC_SPECTACLE.RecupNumAttestInit;
var Q : TQuery;
begin
     if Action <> 'CREATION' then exit;
     Q := OpenSQL('SELECT PGA_NUMATTEST FROM PGAEM WHERE PGA_DATEDEBUT="'+UsDateTime(StrToDate(GetControlText('EDATEDEB')))+'" AND '+
     'PGA_SALARIE="'+LeSalarie+'" ORDER BY PGA_ANNEE,PGA_MOIS',True);
     If not Q.Eof then
     begin
          Q.First;
          SetControlText('NUMATTESTINIT',Q.FindField('PGA_NUMATTEST').AsString);
          SetControlChecked('ATTESTCOMPL',True);
          SetControlChecked('CATTESTINIT',false);
     end;
     ferme(Q);

end;

{PT28 ====>
//PT26 Ajout procedure RecupDerNumObjet ============================
Procedure TOF_PGASSEDIC_SPECTACLE.RecupDerNumObjet;
var Q       : TQuery;
    N_Objet : string;
begin
    if Action <> 'CREATION' then exit;
    N_Objet := '';
    Q := OpenSQL('SELECT PGA_NUMOBJETAEM FROM PGAEM WHERE PGA_ANNEE="'+GetControlText('ANNEE')+
                 '" AND PGA_MOIS="'+GetControlText('MOIS')+'" AND PGA_NUMOBJETAEM <> "" ORDER BY PGA_DATEATTEST DESC',True);

    If not Q.Eof then
    begin
       Q.First;
       N_Objet := Trim(Q.FindField('PGA_NUMOBJETAEM').AsString);
    End;
    ferme(Q);
    SetControlText('NUMOBJET',N_Objet);
end;
<=== PT28}

Initialization
  registerclasses ( [ TOF_PGASSEDIC_SPECTACLE] ) ;
end.

