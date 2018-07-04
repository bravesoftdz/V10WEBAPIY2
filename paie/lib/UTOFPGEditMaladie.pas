{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Attestation maladie
Mots clefs ... : PAIE;MALADIE
*****************************************************************
PT1 13/12/2001 SB V563 Fiche de bug n°392
                       Ré écriture de la procedure AffectSalaire
PT2-1 14/02/2002 SB V571 Ajout du motif Paternité
PT2-2 14/02/2002 SB V571 Fiche de bug n°10024
                         Si DateArret fin de mois alors récup du mois d'arrêt
PT2-3 13/03/2002 SB V571 Fiche de bug n°10024
                         Date arrêt effective = Date arrêt + 1 jours
PT3   10/12/2002 SB V591 FQ 10361 modification nom patronymique
PT4   11/12/2002 SB V591 FQ 10311 Reprise des dates d'absence
PT5   20/01/2003 PH V591 Correction gestion des dates sous ORACLE
PT6   21/01/2003 SB V591 FQ 10434 Utilisation de tob pour maj base
PT7   29/01/2003 SB V591 FQ 10475 Affichage du brut abbattu et non du Brut en colonne 3
PT8-1 24/07/2003 SB V_42 FQ 10723 Si date arrête fin de mois alors non reprise du mois en cours
PT8-2 05/08/2003 SB V_42 FQ 10745 Montant du salaire reprise de la base SS et non du Brut abbattu
PT9   15/09/2003 SB V_42 FQ 10786 Erreur utilsation champ BanqueCP
PT10  09/10/2003 SB V_42 FQ 10881 Stockage de la date de naissance
PT11-1 29/10/2003 SB V_42 FQ 10887 Date de naissance non saisissable pour maladie
PT11-2 29/10/2003 SB V_42 Dysfonctionnement MAJ DONNEE en modification
PT12   03/05/2004 SB V_50 FQ 10812 Intégration de la gestion des déclarants
PT13   09/06/2004 MF V_50 Le paramétre TypeAbs permet de cocher la case Maladie
                          ou maternité ou paternité - Qd il y a subrogation (champ
                          ETB_SUBROGATION) contrôle des champ correspondant.
PT14   05/07/2005 SB V_65 FQ 12378 Contrôle avant impression;
PT15   05/07/2005 SB V_65 FQ 12324 Ajout domiciliation de la banque
PT16   18/10/2005 SB V_65 FQ 12405 Ajout affectation périodicité montant cotisation 200 ou 800 H
PT17   09/11/2006 SB V_70 FQ 13504 Anomalie maj zone numerique
PT18  : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT19  : 03/01/2008 FC V_81 FQ 13378 Alimentation automatique des Colonnes 6 7 8 9
PT20 : 10/09/2008 JS Gestion des dates dans attestation maladie / AT FQ 14874
}

unit UTOFPGEditMaladie;

interface
uses StdCtrls, Controls, Classes,sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  UtileAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  HPdfviewer, Vierge,
  Utob,
  PgOutilsTreso, Commun;

type
  TOF_PGMALADIE = class(TOF)
  private
    Loaded: Boolean;
    Matricule,GblEtab,TypeAbs: string;    // PT13
    subrogation : boolean; // PT13
    Mode: string;
    procedure Imprimer(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure AffectSalaire(Sender: TObject);
    procedure AffectDateArret(Sender: TObject);
    procedure UpdateTable(Sender: TObject);
    procedure GestionCas;
    procedure GestionMotif;
    procedure GestionSubrogation(Sender: TObject);
    procedure MiseABlanc(index: integer);
    procedure InitSalarie(Arguments: string);
    procedure InitAttes(Arguments: string);
    procedure InitTaille(ControlParent: TWinControl);
    {       Function  RemplaceVirguleParPoint (St : String) : String; PT6 }
    procedure AfficheCheck(Sender: TObject);
    procedure AffichDeclarant (Sender: TObject);  { PT12 }
    Function  SaisieNonValide : Boolean; { PT14 }
    procedure Ck200Click (Sender: TObject);  { PT16  13/09/2005 }
    procedure Ck800Click (Sender: TObject);  { PT16  13/09/2005 }
    procedure MotifAbsClick (Sender: TObject);  //PT19

  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnNew; override;
  end;
implementation

uses PgEditOutils, EntPaie, DB;


{ TOF_PGMALADIE }

procedure TOF_PGMALADIE.OnArgument(Arguments: string);
var
  Edit                       : THEdit;
  NumEdit                    : THNumEdit;
  i                          : integer;
  CasGen, CasPart, CkMotif,ck: TCheckBox;
  PDFMaladie                 : TPDFPreview;
  DateAbsence                : TDateTime;
  St                         : String;
  QAttes                     : TQuery;
begin
  inherited;
  MiseABlanc(1);
  //debut PT20
  SetControlText ('DATEDERTRAV', DateToStr (idate1900));
  SetControlText ('DATEREPTRAV', DateToStr (idate1900));
  SetControlText ('DATEARRET', DateToStr (idate1900));
  SetControlText ('DATEDEBPER', DateToStr (idate1900));
  SetControlText ('DATEFINPER', DateToStr (idate1900));
  SetControlText ('DATEDEBPER1', DateToStr (idate1900));
  SetControlText ('DATEDEBPER2', DateToStr (idate1900));
  SetControlText ('DATEDEBPER3', DateToStr (idate1900));
  SetControlText ('DATEDEBPER4', DateToStr (idate1900));
  SetControlText ('DATEDEBPER5', DateToStr (idate1900));
  SetControlText ('DATEDEBPER6', DateToStr (idate1900));
  SetControlText ('DATEDEBPER7', DateToStr (idate1900));
  SetControlText ('DATEDEBPER8', DateToStr (idate1900));
  SetControlText ('DATEDEBPER9', DateToStr (idate1900));
  SetControlText ('DATEDEBPER10', DateToStr (idate1900));
  SetControlText ('DATEDEBPER11', DateToStr (idate1900));
  SetControlText ('DATEDEBPER12', DateToStr (idate1900));
  SetControlText ('DATEFINPER1', DateToStr (idate1900));
  SetControlText ('DATEFINPER2', DateToStr (idate1900));
  SetControlText ('DATEFINPER3', DateToStr (idate1900));
  SetControlText ('DATEFINPER4', DateToStr (idate1900));
  SetControlText ('DATEFINPER5', DateToStr (idate1900));
  SetControlText ('DATEFINPER6', DateToStr (idate1900));
  SetControlText ('DATEFINPER7', DateToStr (idate1900));
  SetControlText ('DATEFINPER8', DateToStr (idate1900));
  SetControlText ('DATEFINPER9', DateToStr (idate1900));
  SetControlText ('DATEFINPER10', DateToStr (idate1900));
  SetControlText ('DATEFINPER11', DateToStr (idate1900));
  SetControlText ('DATEFINPER12', DateToStr (idate1900));
  SetControlText ('DATEDEBSUBR', DateToStr (idate1900));
  SetControlText ('DATEFINSUBR', DateToStr (idate1900));
  SetControlText ('DATENAISSANCE', DateToStr (idate1900));
  // fin PT20

  Matricule := ReadTokenSt(Arguments);
  Mode := ReadTokenSt(Arguments);
  TypeAbs := ReadTokenSt(Arguments);      // PT13
  Arguments := Trim(Arguments);

  //DEB PT4
  if IsValidDate(Arguments) then DateAbsence := StrToDate(Arguments) else DateAbsence := idate1900;
  if Matricule <> '' then InitSalarie(Matricule);
  if DateAbsence > idate1900 then
    SetControlText('DATEARRET', DateToStr(DateAbsence));
  //FIN PT4

  { DEB PT12 }
  St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GblEtab+'%") '+
      ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MAL%" )  ' ;
  SetControlProperty('DECLARANT', 'Plus' ,St );
  Edit:=THEdit(GetControl('DECLARANT'));
  if Edit<>nil then Edit.OnExit := AffichDeclarant;
  { FIN PT12 }

{
  PDFMaladie := TPDFPreview(GetControl('PDFMALADIE'));
  if PDFMaladie <> nil then
  begin
    PDFMaladie.PDFPath := VH_Paie.PGCheminRech + '\MALADIE.pdf'; //Affectation du fichier à recup
  end;
}
  TFVierge(Ecran).BImprimer.OnClick := Imprimer;
  TFVierge(Ecran).BValider.OnClick := UpdateTable;

  //Evenements
  Edit := THEdit(GetControl('DATEARRET'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.OnExit := AffectSalaire;
  end;
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.OnExit := AffectDateArret;
  end;
  Edit := THEdit(GetControl('DATEREPTRAV'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
  end;
  Edit := THEdit(GetControl('DATEDEBPER'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
  end;
  Edit := THEdit(GetControl('DATEFINPER'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
  end;

  for i := 1 to 12 do
  begin
    Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
    if Edit <> nil then
    begin
      Edit.OnDblClick := DateElipsisclick;
    end;
    Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
    if Edit <> nil then
    begin
      Edit.OnDblClick := DateElipsisclick;
    end;
    NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('MTSALRET' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;

    //DEB PT19
    Edit := THEdit(GetControl('MOTIFABS' + IntToStr(i)));
    if Edit <> nil then
      Edit.OnExit := MotifAbsClick;
    //FIN PT19
  end;

  CkMotif := TCheckBox(GetControl('CKMOTIFMAL'));
  if CkMotif <> nil then CkMotif.OnClick := AffectSalaire;
  CkMotif := TCheckBox(GetControl('CKMOTIFMAT'));
  if CkMotif <> nil then CkMotif.OnClick := AffectSalaire;
  CkMotif := TCheckBox(GetControl('CKMOTIFPAT')); //PT2-1
  if CkMotif <> nil then CkMotif.OnClick := AffectSalaire;

  Edit := THEdit(GetControl('DATEDEBSUBR'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Edit := THEdit(GetControl('DATEFINSUBR'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Edit := THEdit(GetControl('DENOMINATION'));
  if Edit <> nil then
  begin
// PT13    Edit.OnDblClick := DateElipsisclick;
    Edit.OnExit := GestionSubrogation;
  end;
  Edit := THEdit(GetControl('DATEFAIT'));
  if Edit <> nil then
  begin
    Edit.text := DateToStr(Date);
    Edit.OnDblClick := DateElipsisclick;
  end;
  Edit := THEdit(GetControl('DATENAISSANCE')); //PT10
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;

  CasGen := TCheckBox(GetControl('CKCASGEN'));
  if CasGen <> nil then CasGen.OnClick := AffectSalaire;
  CasPart := TCheckBox(GetControl('CKCASPART'));
  if CasPart <> nil then CasPart.OnClick := AffectSalaire;

  if (ExisteSQL('SELECT PAS_SALARIE FROM ATTESTATIONS WHERE PAS_SALARIE="' + Matricule + '"') = TRUE) and (Mode = 'MODIFICATION') then
    InitAttes(Matricule)
  else
    Begin { DEB PT12 }
    QAttes:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                    'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GblEtab+'%") '+
                    'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MAL%" )  '+
                    'ORDER BY PDA_ETABLISSEMENT DESC',True);
    If Not QAttes.eof then
      Begin
      SetControlText('DECLARANT' ,QAttes.FindField('PDA_DECLARANTATTES').AsString);
      AffichDeclarant(nil);
      End;
    Ferme(QAttes);
    End;  { FIN PT12 }

  SetControlVisible('PAIEEURO', False);

  { DEB PT16 13/09/2005 }
  ck := TCheckBox(GetControl('CK200'));
  if assigned(ck) then CK.OnClick := Ck200Click;

  ck := TCheckBox(GetControl('CK800'));
  if assigned(ck) then CK.OnClick := Ck800Click;
  { FIN PT16 13/09/2005 }

SetPlusBanqueCP (GetControl ('DENOMINATION'));  //PT18
end;

procedure TOF_PGMALADIE.InitSalarie(Arguments: string);
var
  QSal: TQuery;
  Defaut: THLabel;
  Check : TCheckBox;
begin
  subrogation := False;
  QSal := OpenSql('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,' +
    'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,PSA_SEXE,PSA_NOMJF,' +
    'ET_ETABLISSEMENT,' +
    'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,' +
    'ET_TELEPHONE,ET_SIRET,ET_FAX ' +
    ',ETB_SUBROGATION '+                         // PT13
    'FROM SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT '+ // PT13
    'WHERE PSA_SALARIE="' + Arguments + '" ', TRUE);
  if not QSal.EOF then //PORTAGECWAS
  begin
    GblEtab := QSal.FindField('ET_ETABLISSEMENT').asstring; { PT12 }
    Defaut := THLabel(GetControl('ETLIBELLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_LIBELLE').asstring;
    Defaut := THLabel(GetControl('ETADRESSE1'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE1').asstring;
    Defaut := THLabel(GetControl('ETADRESSE23'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE2').asstring + ' ' + QSal.FindField('ET_ADRESSE3').asstring;

    Defaut := THLabel(GetControl('ETCPVILLE'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring, 'STR', 5) + '    ' + QSal.FindField('ET_VILLE').asstring;
    Defaut := THLabel(GetControl('ETTELEPHONE'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_TELEPHONE').asstring, 'STR', 2);
    Defaut := THLabel(GetControl('ETSIRET'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_SIRET').asstring, 'STR', 2);

    Defaut := THLabel(GetControl('PSASALARIE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_SALARIE').asstring;
    { DEB PT3 Reprise du nom de jeune fille pour les filles
      Defaut := THLabel(GetControl('PSALIBELLE'));
      if defaut <>nil then  Defaut.caption:=QSal.FindField('PSA_LIBELLE').asstring+'   '+QSal.FindField('PSA_PRENOM').asstring;}
    if QSal.FindField('PSA_SEXE').asstring = 'M' then
      SetControlText('PSALIBELLE', QSal.FindField('PSA_LIBELLE').asstring + '   ' + QSal.FindField('PSA_PRENOM').asstring)
    else
      if QSal.FindField('PSA_NOMJF').asstring <> '' then
      SetControlText('PSALIBELLE', QSal.FindField('PSA_NOMJF').asstring + '   ' + QSal.FindField('PSA_PRENOM').asstring + ' (épouse ' + QSal.FindField('PSA_LIBELLE').asstring +
        ')')
    else
      SetControlText('PSALIBELLE', QSal.FindField('PSA_LIBELLE').asstring + '   ' + QSal.FindField('PSA_PRENOM').asstring);
    {FIN PT3}
    Defaut := THLabel(GetControl('PSAADRESSE1'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_ADRESSE1').asstring;
    Defaut := THLabel(GetControl('PSAADRESSE23'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_ADRESSE2').asstring + '  ' + QSal.FindField('PSA_ADRESSE3').asstring;
    Defaut := THLabel(GetControl('PSAEMPLOI'));
    if defaut <> nil then Defaut.caption := RechDom('PGLIBEMPLOI', QSal.FindField('PSA_LIBELLEEMPLOI').asstring, False);

    Defaut := THLabel(GetControl('PSACPVILLE'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring, 'STR', 5) + '    ' + QSal.FindField('PSA_VILLE').asstring;
    Defaut := THLabel(GetControl('PSASECU1'));
    if (defaut <> nil) then
      Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 1, 13), 'STR', 2);
    Defaut := THLabel(GetControl('PSASECU2'));
    if (defaut <> nil) then
      Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 14, 2), 'STR', 2);
// d PT13
    subrogation := QSal.FindField('ETB_SUBROGATION').Asstring = 'X';
// f PT13
  end;
  Ferme(QSal);

// d PT13
  Check := TcheckBox(GetControl('CKMOTIFMAL'));
  if check <> nil then
      if TypeAbs = 'MAN' then Check.Checked := True;
  Check := TcheckBox(GetControl('CKMOTIFMAT'));
  if check <> nil then
      if TypeAbs = 'MAT' then Check.Checked := True;
  Check := TcheckBox(GetControl('CKMOTIFPAT'));
  if check <> nil then
      if TypeAbs = 'PAT' then Check.Checked := True;
// f PT13

end;


procedure TOF_PGMALADIE.InitAttes(Arguments: string);
var
  QAttes: TQuery;
  Check: TcheckBox;
  Edit: THEdit;
  NumEdit: THNumEdit;
  libelle, cas: string;
  i: integer;
begin
  Loaded := True;
  QAttes := OpenSql('SELECT PAS_TRAVAILTEMP,PAS_DERNIERJOUR,PAS_DATEARRET,PAS_REPRISEARRET,' +
    'PAS_MOTIFARRET,PAS_SITUATION,PAS_NONREPRIS,PAS_REPRISEPARTIEL,PAS_CASGEN,' +
    'PAS_MONTANT,PAS_PLUSDE,PAS_PERIODEDEBUT,PAS_PERIODEFIN,PAS_DECLARLIEU,' +
    'PAS_DECLARDATE,PAS_DECLARPERS,PAS_DECLARQUAL,PAS_SUBDEBUT,PAS_SUBFIN,' +
    'PAS_SUBCOMPTE,PAS_SUBCPTEINT,PAS_SUBMONNAIE,PAS_DATEADHESION FROM  ATTESTATIONS ' +
    'WHERE PAS_SALARIE="' + Arguments + '" AND PAS_TYPEATTEST="MAL" ', True);
  if not QAttes.EOF then //PORTAGECWAS
  begin
    Check := TcheckBox(GetControl('ETTEMPOTRAVAIL'));
    if check <> nil then Check.Checked := (QAttes.FindField('PAS_TRAVAILTEMP').asstring = 'X');
    if (QAttes.FindField('PAS_DERNIERJOUR').AsDateTime <> IDate1900) then
      SetControlText('DATEDERTRAV', DateToStr(QAttes.FindField('PAS_DERNIERJOUR').AsDateTime));
    if (QAttes.FindField('PAS_DATEARRET').AsDateTime <> IDate1900) then
      SetControlText('DATEARRET', DateToStr(QAttes.FindField('PAS_DATEARRET').AsDateTime));
    SetControlText('DATEREPTRAV', QAttes.FindField('PAS_REPRISEARRET').asstring);
    Check := TcheckBox(GetControl('CKMOTIFMAL'));
    if check <> nil then
      if QAttes.FindField('PAS_MOTIFARRET').asstring = 'MAL' then Check.Checked := True;
    Check := TcheckBox(GetControl('CKMOTIFMAT'));
    if check <> nil then
      if QAttes.FindField('PAS_MOTIFARRET').asstring = 'MAT' then Check.Checked := True;
    Check := TcheckBox(GetControl('CKMOTIFPAT')); //PT2-1
    if check <> nil then
      if QAttes.FindField('PAS_MOTIFARRET').asstring = 'PAT' then Check.Checked := True;
    GestionMotif;

    SetControlText('SITUATIONARRET', QAttes.FindField('PAS_SITUATION').asstring);
    Check := TcheckBox(GetControl('CKNONREPRIS'));
    if check <> nil then Check.Checked := (QAttes.FindField('PAS_SITUATION').asstring = 'X');
    Check := TcheckBox(GetControl('CKMOTIFPER'));
    if check <> nil then
      if QAttes.FindField('PAS_REPRISEPARTIEL').asstring = 'PER' then Check.Checked := True;
    Check := TcheckBox(GetControl('CKMOTIFMED'));
    if check <> nil then
      if QAttes.FindField('PAS_REPRISEPARTIEL').asstring = 'MED' then Check.Checked := True;
    Check := TcheckBox(GetControl('CKCASGEN'));
    if check <> nil then
      if QAttes.FindField('PAS_CASGEN').asstring = 'GEN' then Check.Checked := True;
    Check := TcheckBox(GetControl('CKCASPART'));
    if check <> nil then
      if QAttes.FindField('PAS_CASGEN').asstring = 'PAR' then Check.Checked := True;
    libelle := '';
    cas := QAttes.FindField('PAS_CASGEN').asstring;
    if cas = 'GEN' then libelle := 'MTCOT200' else if cas = 'PAR' then libelle := 'MTCOT800';
    if libelle <> '' then
    begin
      NumEdit := THNumEdit(GetControl(Libelle));
      if (QAttes.FindField('PAS_MONTANT').asFloat <> 0) and (NumEdit <> nil) then
        NumEdit.Value := QAttes.FindField('PAS_MONTANT').asFloat;
    end;
    if cas = 'GEN' then libelle := 'CK200' else if cas = 'PAR' then libelle := 'CK800';
    if libelle <> '' then Check := TcheckBox(GetControl(libelle));
    if check <> nil then Check.Checked := (QAttes.FindField('PAS_PLUSDE').asstring = 'X');
    if (QAttes.FindField('PAS_PERIODEDEBUT').AsDateTime <> IDate1900) then
      SetControlText('DATEDEBPER', DateToStr(QAttes.FindField('PAS_PERIODEDEBUT').AsDateTime));
    if (QAttes.FindField('PAS_PERIODEFIN').AsDateTime <> IDate1900) then
      SetControlText('DATEFINPER', DateToStr(QAttes.FindField('PAS_PERIODEFIN').AsDateTime));
    SetControlText('LIEUFAIT', QAttes.FindField('PAS_DECLARLIEU').asstring);
    if (QAttes.FindField('PAS_DECLARDATE').AsDateTime <> IDate1900) then
      SetControlText('DATEFAIT', DateToStr(QAttes.FindField('PAS_DECLARDATE').AsDateTime));
    SetControlText('NOMSIGNE', QAttes.FindField('PAS_DECLARPERS').asstring);
    SetControlText('QUALITESIGNE', QAttes.FindField('PAS_DECLARQUAL').asstring);
    if (QAttes.FindField('PAS_SUBDEBUT').AsDateTime <> IDate1900) then
      SetControlText('DATEDEBSUBR', DateToStr(QAttes.FindField('PAS_SUBDEBUT').AsDateTime));
    if (QAttes.FindField('PAS_SUBFIN').AsDateTime <> IDate1900) then
      SetControlText('DATEFINSUBR', DateToStr(QAttes.FindField('PAS_SUBFIN').AsDateTime));
    SetControlText('DENOMINATION', QAttes.FindField('PAS_SUBCOMPTE').asstring);
    SetControlText('INTITULE', QAttes.FindField('PAS_SUBCPTEINT').asstring);
    Check := TcheckBox(GetControl('PAIEEURO'));
    if check <> nil then Check.Checked := (QAttes.FindField('PAS_SUBMONNAIE').AsString = 'X');
    if (QAttes.FindField('PAS_DATEADHESION').AsDateTime <> IDate1900) then //PT10
      SetControlText('DATENAISSANCE', DateToStr(QAttes.FindField('PAS_DATEADHESION').AsDateTime));
  end;
  Ferme(QAttes);
  //Init Tableau
  QAttes := OpenSql('SELECT PAL_SALARIE,PAL_TYPEATTEST,PAL_ORDRE,PAL_MOIS,' +
    'PAL_DATEDEBUT,PAL_DATEFIN,PAL_SALAIRE,PAL_MONNAIESAL,PAL_REGCOTPER1,PAL_REGCOTPER2,' +
    'PAL_MOTIF,PAL_NBHEURES,PAL_NBHEURESCOMPL,PAL_SALAIRERET,PAL_MONNAIESALRET ' +
    'FROM ATTSALAIRES WHERE PAL_SALARIE="' + Arguments + '" AND PAL_TYPEATTEST="MAL" ' +
    'ORDER BY PAL_MOIS', True);
  i := 0;
  while not QAttes.EOF do
  begin //B While
    i := i + 1;
    Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
    if (Edit <> nil) and (QAttes.FindField('PAL_DATEDEBUT').AsDateTime <> IDate1900) then
      Edit.Text := DateToStr(QAttes.FindField('PAL_DATEDEBUT').AsDateTime);
    Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
    if (Edit <> nil) and (QAttes.FindField('PAL_DATEFIN').AsDateTime <> IDate1900) then
      Edit.text := DateToStr(QAttes.FindField('PAL_DATEFIN').AsDateTime);
    NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_SALAIRE').asFloat;
    Check := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
    if Check <> nil then
    begin
      Check.checked := (QAttes.FindField('PAL_MONNAIESAL').AsString = 'X');
      Check.Visible := (QAttes.FindField('PAL_SALAIRE').asFloat <> 0);
    end;
    NumEdit := THNumEdit(GetControl('MTREGPER1' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_REGCOTPER1').asFloat;
    NumEdit := THNumEdit(GetControl('MTREGPER2' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_REGCOTPER2').asFloat;
    Edit := THEdit(GetControl('MOTIFABS' + IntToStr(i)));
    if Edit <> nil then Edit.text := QAttes.FindField('PAL_MOTIF').asstring;
    Edit := THEdit(GetControl('NBHEURREEL' + IntToStr(i)));
    if Edit <> nil then Edit.text := IntToStr(QAttes.FindField('PAL_NBHEURES').AsInteger); // DB2
    Edit := THEdit(GetControl('NBHEURCOMPL' + IntToStr(i)));
    if Edit <> nil then Edit.text := IntToStr(QAttes.FindField('PAL_NBHEURESCOMPL').AsInteger); // DB2
    NumEdit := THNumEdit(GetControl('MTSALRET' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_SALAIRERET').asFloat;
    Check := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(i)));
    if Check <> nil then
    begin
      Check.checked := (QAttes.FindField('PAL_MONNAIESALRET').asstring = 'X');
      Check.Visible := (QAttes.FindField('PAL_SALAIRERET').asFloat <> 0);
    end;
    QAttes.next;
  end; //E While
  Ferme(QAttes);
  Loaded := False;
end;

procedure TOF_PGMALADIE.Imprimer(Sender: TObject);
var
  Pages: TPageControl;
  {$IFDEF EAGLCLIENT}
  StPages : String;
  {$ENDIF}
begin
  if SaisieNonValide then Exit;
  Pages := TPageControl(GetControl('PAGES'));
  if Pages <> nil then
    {$IFNDEF EAGLCLIENT}
    LanceEtat('E', 'PAT', 'MAL', True, False, False, Pages, '', '', False);
  {$ELSE}
     StPages := AglGetCriteres(Pages, FALSE);
  LanceEtat('E', 'PAT', 'MAL', True, False, False, nil, '', '', False, 0, StPages);
  {$ENDIF}
end;

procedure TOF_PGMALADIE.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGMALADIE.AffectSalaire(Sender: TObject);
var
  QRech: TQuery;
  Edit, Zone: THedit;
  DateD, DerJour,DateDeb, DateFin : TDateTime; {Gestion de la date de passage à l'euro}
  i, j: integer;
  MontZone: THNumEdit;
  CasGen, CasPart, MotifMal, MotifMat, CkEuro, Ck, MotifPat: TCheckBox;
  Tob_Montant, TMont: Tob;
  Trouve: boolean;
begin
  if Loaded = True then exit;
  MontZone := nil;
  Ck := nil;
  GestionCas;
  GestionMotif;
  Edit := THEdit(GetControl('DATEARRET'));
  CasGen := TCheckBox(GetControl('CKCASGEN'));
  CasPart := TCheckBox(GetControl('CKCASPART'));
  MotifMal := TCheckBox(GetControl('CKMOTIFMAL'));
  MotifMat := TCheckBox(GetControl('CKMOTIFMAT'));
  MotifPat := TCheckBox(GetControl('CKMOTIFPAT')); //PT2-1
  // deb PT20
  //SetControlText('DATEDEBPER','');   { PT16 }
  //SetControlText('DATEFINPER','');   { PT16 }
  //fin PT20
  if (Edit <> nil) and (CasGen <> nil) and (CasPart <> nil) and (MotifMal <> nil) and (MotifMat <> nil) then
  begin
    if (Edit.text = '') and (CasGen.checked = False) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text <> '') and (CasGen.checked = False) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text <> '') and (CasGen.checked = True) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text <> '') and (CasGen.checked = False) and (CasPart.checked = True) and (MotifMal.checked = False) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text <> '') and (CasGen.checked = False) and (CasPart.checked = False) and (MotifMal.checked = True) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text <> '') and (CasGen.checked = False) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = True) then MiseABlanc(1);

    if (Edit.text = '') and (CasGen.checked = True) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = False) and (CasPart.checked = True) and (MotifMal.checked = False) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = False) and (CasPart.checked = False) and (MotifMal.checked = True) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = False) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = True) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = True) and (CasPart.checked = False) and (MotifMal.checked = True) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = True) and (CasPart.checked = False) and (MotifMal.checked = False) and (MotifMat.Checked = True) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = False) and (CasPart.checked = True) and (MotifMal.checked = True) and (MotifMat.Checked = False) then MiseABlanc(1);
    if (Edit.text = '') and (CasGen.checked = False) and (CasPart.checked = True) and (MotifMal.checked = False) and (MotifMat.Checked = True) then MiseABlanc(1);
  end;

  //DEB PT1
  if (Edit <> nil) and (CasGen <> nil) and (CasPart <> nil) and (MotifMal <> nil) and (MotifMat <> nil) and (MotifPat <> nil) then //PT2-1
  begin
    //Cas General  et cas particulier
    if (Edit.Text <> '') and ((CasGen.Checked = True) or (CasPart.Checked = True)) and ((MotifMal.Checked = True) or (MotifMat.Checked = True) or (MotifPat.Checked = True)) then
      //PT2-1
    begin
      DerJour := StrToDate(Edit.text);
      DateFin := DebutdeMois(DerJour)-1; { PT8-1 }
      {if DerJour=FinDeMois(DerJour) then DecalUnMois:=1 else DecalUnMois:=0; //PT2-2 PT8-1 Mise en commentaire }
      //Affectation "Montant de la cotisation"
      if (CasGen.Checked = True) then
      begin
        MontZone := THNumEdit(GetControl('MTCOT200'));
        DateDeb := DebutDeMois(PlusMois(DebutDeMois(DerJour)-1, -5)); { PT8-1  Suppr DecalUnMois }
      end
      else
        if (CasPart.Checked = True) then
      begin
        MontZone := THNumEdit(GetControl('MTCOT800'));
        DateDeb := DebutDeMois(PlusMois(DebutDeMois(DerJour)-1, -11)); { PT8-1  Suppr DecalUnMois }
      end;
      if MontZone <> nil then
      begin
        QRech := OpenSql('SELECT SUM(PHB_MTSALARIAL) AS CRDS FROM HISTOBULLETIN ' +
          'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' + //**//
          'WHERE PCT_SOUMISMALAD="X"  AND PHB_SALARIE="' + Matricule + '"  ' +
          'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ', True); //PT2-2 Ajout du <=
        if not QRech.EOF then //PORTAGECWAS
          Begin
          MontZone.Value := QRech.FindField('CRDS').asFloat;
          SetControlText('DATEDEBPER',DateToStr(DateDeb)); { PT16 }
          SetControlText('DATEFINPER',DateToStr(DateFin)); { PT16 }
          End;
        Ferme(Qrech);
      end;
      //Affectation "PLus de 200h"
      if (CasGen.Checked = True) then
      begin
      Ck := TCheckBox(GetControl('CK200'));
      DateDeb := DebutDeMois(PlusMois(DerJour, -2)); { PT8-1  Suppr DecalUnMois }
      end
      else
        if (CasPart.Checked = True) then
      begin
        Ck := TCheckBox(GetControl('CK800'));
        DateDeb := DebutDeMois(PlusMois(DerJour, -11)); { PT8-1  Suppr DecalUnMois }
      end;
      if Ck <> nil then
      begin
        QRech := OpenSql('SELECT SUM(PPU_CHEURESTRAV) AS HEURE FROM PAIEENCOURS ' +
          'WHERE PPU_SALARIE="' + Matricule + '"  ' +
          'AND PPU_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ', True); //PT2-2  Ajout du <=
        if not QRech.EOF then //PORTAGECWAS
        begin
          if (CasGen.Checked = True) then
            Ck.Checked := (QRech.FindField('HEURE').asFloat >= 200)
          else
            if (CasPart.Checked = True) then
            Ck.Checked := (QRech.FindField('HEURE').asFloat >= 800);
        end;
        Ferme(Qrech);
      end;
      //Remplissage du tableau
      Tob_Montant := Tob.create('Histo_Bulletin', nil, -1);
      if (CasGen.Checked = True) then
        DateDeb := DebutDeMois(PlusMois(DerJour, -3)) { PT8-1  Suppr DecalUnMois }
      else
        if (CasPart.Checked = True) then
        DateDeb := DebutDeMois(PlusMois(DerJour, -12)); { PT8-1  Suppr DecalUnMois }
      {  PT8-2 Mise en commentaire Modification de la requête
       QRech:=OpenSql('SELECT DISTINCT PPU_DATEDEBUT DATEDEBUT,PPU_DATEFIN DATEFIN,'+
      // 'SUM(PPU_CBRUT) AS CBRUT,SUM(PPU_OCBRUT) AS OCBRUT,'+PT7 Mise en commentaire
         'SUM(PPU_CBRUTFISCAL) AS CBRUTFISCAL,'+
         'SUM(PPU_OCBRUTFISCAL) AS OCBRUTFISCAL,'+
         'SUM(PPU_CCOUTSALARIE) AS CCOUTSALARIE,'+
         'SUM(PPU_OCCOUTSALARIE) AS OCCOUTSALARIE '+
         'FROM PAIEENCOURS '+
         'WHERE  PPU_SALARIE="'+Matricule+'" '+
         'AND PPU_DATEDEBUT>="'+UsDateDeb+'" AND PPU_DATEFIN<="'+UsDateFin+'" '+      //PT2-2  Ajout du <=
         'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN ORDER BY PPU_DATEDEBUT,PPU_DATEFIN',TRUE);}
      QRech := OpenSql('SELECT DISTINCT PPU_DATEDEBUT DATEDEBUT,PPU_DATEFIN DATEFIN,' +
        'SUM(PPU_CBASESS) AS CBASESS,' +
        'SUM(PPU_OCBASESS) AS OCBASESS,' +
        'SUM(PPU_CCOUTSALARIE) AS CCOUTSALARIE,' +
        'SUM(PPU_OCCOUTSALARIE) AS OCCOUTSALARIE ' +
        'FROM PAIEENCOURS ' +
        'WHERE  PPU_SALARIE="' + Matricule + '" ' +
        'AND PPU_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ' + //PT2-2  Ajout du <=
        'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN ORDER BY PPU_DATEDEBUT,PPU_DATEFIN', TRUE);

      Tob_Montant.LoadDetailDB('Histo_Bulletin', '', '', QRech, False);
      Ferme(QRech);
      TMont := Tob_montant.FindFirst([''], [''], False);
      while TMont <> nil do
      begin
        TMont.AddChampSup('CRDS', False);
        TMont.AddChampSup('OCRDS', False);
        TMont.PutValue('CRDS', 0);
        TMont.PutValue('OCRDS', 0);
        TMont := Tob_montant.FindNext([''], [''], False);
      end;
      QRech := OpenSql('SELECT DISTINCT PHB_DATEDEBUT DATEDEBUT,PHB_DATEFIN DATEFIN,' +
        'SUM(PHB_MTSALARIAL) AS CRDS,' +
        'SUM(PHB_OMTSALARIAL) AS OCRDS FROM HISTOBULLETIN ' +
        'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' + //**//
        'WHERE PCT_ETUDEDROIT="X"  AND PHB_SALARIE="' + Matricule + '" ' +
        'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' + //PT2-2   Ajout du <=
        'GROUP BY PHB_DATEDEBUT,PHB_DATEFIN ORDER BY PHB_DATEDEBUT,PHB_DATEFIN', TRUE);
      Trouve := False;
      while not QRech.eof do
      begin
        TMont := Tob_montant.FindFirst([''], [''], False);
        while TMont <> nil do
        begin
          if (TMont.GetValue('DATEDEBUT') = QRech.FindField('DATEDEBUT').AsDateTime)
            and (TMont.GetValue('DATEFIN') = QRech.FindField('DATEFIN').AsDateTime) then
          begin
            Trouve := True;
            TMont.PutValue('CRDS', TMont.GetValue('CRDS') + QRech.FindField('CRDS').AsFloat);
            TMont.PutValue('OCRDS', TMont.GetValue('OCRDS') + QRech.FindField('OCRDS').AsFloat);
          end;
          TMont := Tob_montant.FindNext([''], [''], False);
        end;
        if not Trouve then
        begin
          TMont := Tob.create('La fille', tob_Montant, -1);
          if TMont <> nil then
          begin
            TMont.PutValue('DATEDEBUT', QRech.FindField('DATEDEBUT').AsDateTime);
            TMont.PutValue('DATEFIN', QRech.FindField('DATEFIN').AsDateTime);
            TMont.PutValue('CRDS', QRech.FindField('CRDS').AsFloat);
            TMont.PutValue('OCRDS', QRech.FindField('OCRDS').AsFloat);
          end;
        end;
        Trouve := False;
        QRech.next;
      end;
      Ferme(QRech);
      j := 0;
      i := 0;
      TMont := Tob_montant.FindFirst([''], [''], False);
      while TMont <> nil do
      begin //B While
        i := i + 1;
        j := j + 1;
        Zone := THEdit(GetControl('DATEDEBPER' + IntToStr(j)));
        if Zone <> nil then Zone.text := DateToStr(TMont.GetValue('DATEDEBUT'));
        Zone := THEdit(GetControl('DATEFINPER' + IntToStr(j)));
        if Zone <> nil then Zone.text := DateToStr(TMont.GetValue('DATEFIN'));
        DateD := TMont.GetValue('DATEDEBUT');
        MontZone := THNumEdit(GetControl('MTSAL' + IntToStr(j)));
        if MontZone <> nil then
        begin
          {Gestion de la date de passage à l'euro
                  if MotifMal.Checked=True Then MontZone.Value:=QRech.FindField('PPU_CBRUT').asFloat;
                  if MotifMat.Checked=True Then MontZone.Value:=QRech.FindField('PPU_CBRUTFISCAL').asFloat-QRech.FindField('PPU_CCOUTSALARIE').asFloat+QRech.FindField('CRDS').asFloat;
                  End;
                CkEuro:=TCheckBox(GetControl('CKSALEURO'+IntToStr(j)));
                CkEuro:=TCheckBox(GetControl('CKSALRETEURO'+IntToStr(j)));
          }
          if (VH_Paie.PGTenueEuro = FALSE) then
          begin
            if MotifMal.Checked = True then
              MontZone.Value := TMont.GetValue('CBASESS'); //PT7 //PT8-2
            if (MotifMat.Checked = True) or (MotifPat.Checked = True) then //PT2-1
              MontZone.Value := TMont.GetValue('CBASESS') - TMont.GetValue('CCOUTSALARIE') + TMont.GetValue('CRDS'); //PT8-2
            CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(j)));
            if CkEuro <> nil then
            begin
              CkEuro.Checked := VH_Paie.PGTenueEuro;
              CkEuro.Visible := True;
            end;
            CkEuro := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(j)));
            if CkEuro <> nil then
            begin
              CkEuro.Checked := VH_Paie.PGTenueEuro;
              CkEuro.Visible := True;
            end;
          end
          else
          begin
            if (VH_Paie.PGDateBasculEuro <= DateD) then
            begin
              if MotifMal.Checked = True then
                MontZone.Value := TMont.GetValue('CBASESS'); //PT7 //PT8-2
              if (MotifMat.Checked = True) or (MotifPat.Checked = True) then //PT2-1
                MontZone.Value := TMont.GetValue('CBASESS') - TMont.GetValue('CCOUTSALARIE') + TMont.GetValue('CRDS'); //PT8-2
              CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(j)));
              if CkEuro <> nil then
              begin
                CkEuro.Checked := VH_Paie.PGTenueEuro;
                CkEuro.Visible := True;
              end;
              CkEuro := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(j)));
              if CkEuro <> nil then
              begin
                CkEuro.Checked := VH_Paie.PGTenueEuro;
                CkEuro.Visible := True;
              end;
            end
            else
            begin
              if MotifMal.Checked = True then
                MontZone.Value := TMont.GetValue('OCBASESS'); //PT7 //PT8-2
              if (MotifMat.Checked = True) or (MotifPat.Checked = True) then //PT2-1
                MontZone.Value := TMont.GetValue('OCBASESS') - TMont.GetValue('OCCOUTSALARIE') + TMont.GetValue('OCRDS'); //PT8-2
              CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(j)));
              if CkEuro <> nil then
              begin
                CkEuro.Checked := False;
                CkEuro.Visible := True;
              end;
              CkEuro := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(j)));
              if CkEuro <> nil then
              begin
                CkEuro.Checked := False;
                CkEuro.Visible := True;
              end;
            end;
          end;
        end;
        {FIN Gestion de la date de passage à l'euro}
        Tmont := Tob_montant.FindNext([''], [''], False); //QRech.Next;
      end; //End While
      if i < 12 then MiseABlanc(i + 1);
      if Tob_Montant <> nil then Tob_Montant.free;
    end; //Fin Cas Général  et cas particulier

    //FIN PT1
  end;
  if GetControlText('CKMOTIFMAL') = 'X' then // SetControlText('DATENAISSANCE', ''); //PT11-1
     SetControlText('DATENAISSANCE', DateToStr (idate1900)); //PT20
  SetControlEnabled('DATENAISSANCE', GetControlText('CKMOTIFMAL') <> 'X');
end;

procedure TOF_PGMALADIE.GestionCas;
var
  MontZone: THNumEdit;
  CasGen, CasPart, Check: TCheckBox;
begin
  //Accessibilité des coches Cas Genéral et Cas Particuliers
  CasGen := TCheckBox(GetControl('CKCASGEN'));
  CasPart := TCheckBox(GetControl('CKCASPART'));
  if (CasGen <> nil) and (CasPart <> nil) then
  begin
    if CasGen.Checked = True then
    begin
      CasPart.enabled := False;
      MontZone := THNumEdit(GetControl('MTCOT200'));
      if MontZone <> nil then MontZone.Enabled := True;
      Check := TCheckBox(GetControl('CK200'));
      if Check <> nil then Check.Enabled := True;
      Check := TCheckBox(GetControl('CKMT200'));
      if Check <> nil then
      begin
        Check.Enabled := True;
        Check.Checked := VH_Paie.PGTenueEuro;
      end;
    end
    else
    begin
      CasPart.enabled := True;
      MontZone := THNumEdit(GetControl('MTCOT200'));
      if MontZone <> nil then
      begin
        MontZone.Enabled := False;
        MontZone.Value := 0;
      end;
      Check := TCheckBox(GetControl('CK200'));
      if Check <> nil then
      begin
        Check.Enabled := False;
        Check.Checked := False;
      end;
      Check := TCheckBox(GetControl('CKMT200'));
      if Check <> nil then
      begin
        Check.Enabled := False;
        Check.Checked := False;
      end;
    end;
    if CasPart.Checked = True then
    begin
      CasGen.enabled := False;
      MontZone := THNumEdit(GetControl('MTCOT800'));
      if MontZone <> nil then MontZone.Enabled := True;
      Check := TCheckBox(GetControl('CK800'));
      if Check <> nil then Check.Enabled := True;
      Check := TCheckBox(GetControl('CKMT800'));
      if Check <> nil then
      begin
        Check.Enabled := True;
        Check.Checked := VH_Paie.PGTenueEuro;
      end;
    end
    else
    begin
      CasGen.enabled := True;
      MontZone := THNumEdit(GetControl('MTCOT800'));
      if MontZone <> nil then
      begin
        MontZone.Enabled := False;
        MontZone.Value := 0;
      end;
      Check := TCheckBox(GetControl('CK800'));
      if Check <> nil then
      begin
        Check.Enabled := False;
        Check.Checked := False;
      end;
      Check := TCheckBox(GetControl('CKMT800'));
      if Check <> nil then
      begin
        Check.Enabled := False;
        Check.Checked := False;
      end;
    end;
  end;
end;


procedure TOF_PGMALADIE.GestionMotif;
var
  CkMotifMal, CkMotifMat, CkMotifPat, Coche: TCheckBox;
begin
  CkMotifMal := TCheckBox(GetControl('CKMOTIFMAL'));
  CkMotifMat := TCheckBox(GetControl('CKMOTIFMAT'));
  CkMotifPat := TCheckBox(GetControl('CKMOTIFPAT'));
  if (CkMotifMal <> nil) and (CkMotifMat <> nil) and (CkMotifPat <> nil) then
  begin //DEB PT2-1
    Coche := nil;
    if CkMotifMal.Checked = True then Coche := TCheckBox(CkMotifMal);
    if CkMotifMat.Checked = True then Coche := TCheckBox(CkMotifMat);
    if CkMotifPat.Checked = True then Coche := TCheckBox(CkMotifPat);
    if Coche <> nil then
    begin
      CkMotifMal.Enabled := (Coche.name = 'CKMOTIFMAL');
      CkMotifMat.Enabled := (Coche.name = 'CKMOTIFMAT');
      CkMotifPat.Enabled := (Coche.name = 'CKMOTIFPAT');
    end
    else
    begin
      CkMotifMal.Enabled := True;
      CkMotifMat.Enabled := True;
      CkMotifPat.Enabled := True;
    end;
    {if CkMotifMal.Checked=True then CkMotifMat.enabled:=False
    else CkMotifMat.enabled:=True;
    if CkMotifMat.Checked=True then CkMotifMal.enabled:=False
    else CkMotifMal.enabled:=True;
             FIN PT2-1}
  end;
end;
procedure TOF_PGMALADIE.MiseABlanc(index: integer);
var
  Zone: THEdit;
  MontZone: THNumEdit;
  i: integer;
  CkEuro: TCheckBox;
begin
  for i := index to 12 do
  begin
   //deb PT20
   // Zone := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
   // if Zone <> nil then
   // begin
   //   Zone.text := '';
   // end;
   // Zone := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
   // if Zone <> nil then
   // begin
   //   Zone.text := '';
   // end;
   //fin PT20
    MontZone := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
    if MontZone <> nil then
    begin
      MontZone.Value := 0;
    end; //MontZone.Visible:=False;
    CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
    if CkEuro <> nil then
    begin
      CkEuro.Checked := False;
    end; //CkEuro.Visible:=False;
    CkEuro.Visible := False;
    MontZone := THNumEdit(GetControl('MTREGPER1' + IntToStr(i)));
    if MontZone <> nil then
    begin
      MontZone.Value := 0;
    end; //MontZone.Visible:=False;
    MontZone := THNumEdit(GetControl('MTREGPER2' + IntToStr(i)));
    if MontZone <> nil then
    begin
      MontZone.Value := 0;
    end; //MontZone.Visible:=False;
    Zone := THEdit(GetControl('MOTIFABS' + IntToStr(i)));
    if Zone <> nil then
    begin
      Zone.text := '';
    end; // Zone.Visible:=False;
    Zone := THEdit(GetControl('NBHEURREEL' + IntToStr(i)));
    if Zone <> nil then
    begin
      Zone.text := '';
    end; //Zone.Visible:=False;
    Zone := THEdit(GetControl('NBHEURCOMPL' + IntToStr(i)));
    if Zone <> nil then
    begin
      Zone.text := '';
    end; //  Zone.Visible:=False;
    MontZone := THNumEdit(GetControl('MTSALRET' + IntToStr(i)));
    if MontZone <> nil then
    begin
      MontZone.Value := 0;
    end; //MontZone.Visible:=False;
    CkEuro := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(i)));
    if CkEuro <> nil then
    begin
      CkEuro.Checked := False;
    end; //CkEuro.Visible:=False;
  end;
end;

procedure TOF_PGMALADIE.GestionSubrogation(Sender: TObject);
var
  deno: THedit;
  QRech: Tquery;
  StPlus : String;
begin
  SetControlEnabled('BANQUE', False);
  SetControlText('BANQUE', '');
  SetControlEnabled('GUICHET', False);
  SetControlText('GUICHET', '');
  SetControlEnabled('COMPTE', False);
  SetControlText('COMPTE', '');
  SetControlEnabled('CLE', False);
  SetControlText('CLE', '');
  SetControlText('INTITULE', '');  { PT15 }
  deno := THedit(GetControl('DENOMINATION'));
  if deno <> nil then
    if deno.text <> '' then
    begin
      StPlus:= PGBanqueCP (True);            //PT18

      QRech := opensql('SELECT BQ_DOMICILIATION,BQ_ETABBQ,BQ_NUMEROCOMPTE,BQ_CLERIB,BQ_GUICHET ' + { PT15 }
        'FROM BANQUECP WHERE BQ_GENERAL="' + deno.text + '"'+StPlus, TRUE); //PT9
      if not QRech.EOF then //PORTAGECWAS
      begin
        SetControlText('INTITULE', QRech.FindField('BQ_DOMICILIATION').asstring); { PT15 }
        SetControlEnabled('BANQUE', True);
        SetControlText('BANQUE', QRech.FindField('BQ_ETABBQ').asstring);
        SetControlEnabled('GUICHET', True);
        SetControlText('GUICHET', QRech.FindField('BQ_GUICHET').asstring);
        SetControlEnabled('COMPTE', True);
        SetControlText('COMPTE', QRech.FindField('BQ_NUMEROCOMPTE').asstring);
        SetControlEnabled('CLE', True);
        SetControlText('CLE', QRech.FindField('BQ_CLERIB').asstring);
      end;
      Ferme(QRech);
    end;
end;


procedure TOF_PGMALADIE.UpdateTable(Sender: TObject);
var
  Check: TcheckBox;
  Edit: THEdit;
  NumEdit: THNumEdit;
  libelle, TravTemp, DerjourTrav, DateRepArret, MotifArret, Situation, NonRepris: string;
  RepPartiel, Cas, Plus, DateDebPer, DateFinPer, LieuDecla, DateDecla, PersDecla, QualDecla: string;
  DateDebSubr, DateFinSubr, CompteBq, Intitule, Monnaie, DateArret: string;
  StSal, MtSalRet, MtRegCotPer1, MtRegCotPer2: string;
  Motif, NbHeures, NbHeureCompl, SalEuro, SalRetEuro, DateNaissance: string;
  Montant: double;
  ordre, i: integer;
  QRech: TQuery;
  Tob_Mere, T: TOB;
begin
  if SaisieNonValide then  Exit;
  // PT5   20/01/2003 PH V591 Correction gestion des dates sous ORACLE
  // Transformation des UsDateTime en DateToStr sur toutes les affectations par
  // defaut de Idate1900
  SetFocusControl('PDFMALADIE');   { PT17 }
  TravTemp := '-';
  ordre := 0;
  Check := TcheckBox(GetControl('ETTEMPOTRAVAIL'));
  if check <> nil then if Check.Checked = True then TravTemp := 'X';
  DerjourTrav := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if edit <> nil then if Edit.text <> '' then DerjourTrav := Edit.text;
  DateArret := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEARRET'));
  if edit <> nil then if Edit.text <> '' then DateArret := Edit.text;
  DateRepArret := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEREPTRAV'));
  if edit <> nil then if Edit.text <> '' then DateRepArret := Edit.text;
  MotifArret := '';
  Check := TcheckBox(GetControl('CKMOTIFMAL'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MAL';
  Check := TcheckBox(GetControl('CKMOTIFMAT'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MAT';
  Check := TcheckBox(GetControl('CKMOTIFPAT')); //PT2-1
  if check <> nil then if Check.Checked = True then MotifArret := 'PAT';
  Situation := '';
  Edit := THEdit(GetControl('SITUATIONARRET'));
  if edit <> nil then if Edit.text <> '' then Situation := Edit.text;
  NonRepris := '-';
  Check := TcheckBox(GetControl('CKNONREPRIS'));
  if check <> nil then if Check.Checked = True then NonRepris := 'X';
  RepPartiel := '';
  Check := TcheckBox(GetControl('CKMOTIFPER'));
  if check <> nil then if Check.Checked = True then RepPartiel := 'PER';
  Check := TcheckBox(GetControl('CKMOTIFMED'));
  if check <> nil then if Check.Checked = True then RepPartiel := 'MED';
  Cas := '';
  Check := TcheckBox(GetControl('CKCASGEN'));
  if check <> nil then if Check.Checked = True then Cas := 'GEN';
  Check := TcheckBox(GetControl('CKCASPART'));
  if check <> nil then if Check.Checked = True then Cas := 'PAR';
  libelle := '';
  Montant := 0;
  if cas = 'GEN' then libelle := 'MTCOT200';
  if cas = 'PAR' then libelle := 'MTCOT800';
  if libelle <> '' then
  begin
    NumEdit := THNumEdit(GetControl(Libelle));
    if NumEdit <> nil then if NumEdit.Value > 0 then Montant := NumEdit.value;
  end;
  Plus := '-';
  if cas = 'GEN' then libelle := 'CK200';
  if cas = 'PAR' then libelle := 'CK800';
  if libelle <> '' then Check := TcheckBox(GetControl(libelle));
  if check <> nil then if Check.Checked = True then Plus := 'X';

  DateDebPer := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDEBPER'));
  if edit <> nil then if Edit.text <> '' then DateDebPer := Edit.text;
  DateFinPer := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFINPER'));
  if edit <> nil then if Edit.text <> '' then DateFinPer := Edit.text;
  LieuDecla := '';
  Edit := THEdit(GetControl('LIEUFAIT'));
  if edit <> nil then if Edit.text <> '' then LieuDecla := Edit.text;
  DateDecla := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFAIT'));
  if edit <> nil then if Edit.text <> '' then DateDecla := Edit.text;
  PersDecla := '';
  Edit := THEdit(GetControl('NOMSIGNE'));
  if edit <> nil then if Edit.text <> '' then PersDecla := Edit.text;
  QualDecla := '';
  Edit := THEdit(GetControl('QUALITESIGNE'));
  if Edit <> nil then if Edit.Text <> '' then QualDecla := Edit.Text;
  DateDebSubr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDEBSUBR'));
  if edit <> nil then if Edit.text <> '' then DateDebSubr := Edit.text;
  DateFinSubr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFINSUBR'));
  if edit <> nil then if Edit.text <> '' then DateFinSubr := Edit.text;
  CompteBq := '';
  Edit := THEdit(GetControl('DENOMINATION'));
  if edit <> nil then if Edit.text <> '' then CompteBq := Edit.text;
  Intitule := '';
  Edit := THEdit(GetControl('INTITULE'));
  if edit <> nil then if Edit.text <> '' then Intitule := Edit.text;
  Monnaie := '-';
  Check := TcheckBox(GetControl('PAIEEURO'));
  if check <> nil then if Check.Checked = True then Monnaie := 'X';
  DateNaissance := DateToStr(idate1900); //PT10
  Edit := THEdit(GetControl('DATENAISSANCE'));
  if edit <> nil then if Edit.text <> '' then DateNaissance := Edit.text;

// d PT13
  if (subrogation = True) and
     ((DateDebSubr = DateToStr(Idate1900)) or
      (DateFinSubr = DateToStr(Idate1900)) or
      (CompteBq = '') or
      (Intitule = '')) then
     PgiBox('Des informations concernant la subrogation sont absentes!','Subrogation');
// f PT13

  {DEB PT6 Recherche de l'ordre sur la table ATTESTATIONS}
  QRech := OpenSql('SELECT PAS_ORDRE FROM ATTESTATIONS ' +
    'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="MAL"', True);
  if not QRech.Eof then ordre := QRech.Fields[0].asinteger;
  Ferme(QRech);
  if ordre = 0 then
  begin
    QRech := OpenSql('SELECT MAX(PAS_ORDRE) FROM ATTESTATIONS ', True);
    if not QRech.EOF then //PORTAGECWAS
      ordre := QRech.Fields[0].asinteger + 1
    else
      ordre := 1;
    Ferme(QRech);
  end;
  {PT6 Utilisation d'une tob pour maj des champs}
  Tob_Mere := Tob.create('ATTESTATIONS', nil, -1);
  T := Tob.create('ATTESTATIONS', Tob_Mere, -1);
  if Ordre > 1 then T.SelectDB('"' + Matricule + '";"MAL";"' + IntToStr(Ordre) + '"', nil); //PT11-2
  if T = nil then exit;
  T.PutValue('PAS_SALARIE', Matricule);
  T.PutValue('PAS_TYPEATTEST', 'MAL');
  T.PutValue('PAS_ORDRE', ordre);
  T.PutValue('PAS_TRAVAILTEMP', TravTemp);
  T.PutValue('PAS_DERNIERJOUR', StrToDate(DerjourTrav));
  T.PutValue('PAS_DATEARRET', StrToDate(DateArret));
  T.PutValue('PAS_REPRISEARRET', StrToDate(DateRepArret));
  T.PutValue('PAS_MOTIFARRET', MotifArret);
  T.PutValue('PAS_SITUATION', Situation);
  T.PutValue('PAS_NONREPRIS', NonRepris);
  T.PutValue('PAS_REPRISEPARTIEL', RepPartiel);
  T.PutValue('PAS_CASGEN', Cas);
  T.PutValue('PAS_MONTANT', FloatToStr(Montant));
  T.PutValue('PAS_PLUSDE', Plus);
  T.PutValue('PAS_PERIODEDEBUT', StrToDate(DateDebPer));
  T.PutValue('PAS_PERIODEFIN', StrToDate(DateFinPer));
  T.PutValue('PAS_DECLARDATE', StrToDate(DateDecla));
  T.PutValue('PAS_DECLARPERS', PersDecla);
  T.PutValue('PAS_DECLARQUAL', QualDecla);
  T.PutValue('PAS_DECLARLIEU', LieuDecla);
  T.PutValue('PAS_SUBDEBUT', StrToDate(DateDebSubr));
  T.PutValue('PAS_SUBFIN', StrToDate(DateFinSubr));
  T.PutValue('PAS_SUBCOMPTE', CompteBq);
  T.PutValue('PAS_SUBCPTEINT', Intitule);
  T.PutValue('PAS_SUBMONNAIE', Monnaie);
  T.PutValue('PAS_DATEADHESION', StrToDate(DateNaissance)); //PT10
  T.InsertOrUpdateDB;
  if Tob_Mere <> nil then Tob_Mere.free;

  {PT6 Mise en commentaire puis supprimé }

  //MAJ Tableau
  Tob_Mere := Tob.create('ATTSALAIRES', nil, -1);

  for i := 1 to 12 do
  begin //B For
    DateDebPer := DateToStr(Idate1900);
    Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateDebPer := Edit.text; //
    DateFinPer := DateToStr(Idate1900);
    Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateFinPer := Edit.text;
    NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
    if NumEdit <> nil then StSal := FloatToStr(NumEdit.value); //FloatToStr(NumEdit.Value);
    Check := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then SalEuro := 'X' else SalEuro := '-';
    MtRegCotPer1 := '0';
    NumEdit := THNumEdit(GetControl('MTREGPER1' + IntToStr(i)));
    if NumEdit <> nil then MtRegCotPer1 := FloatToStr(NumEdit.Value);
    MtRegCotPer2 := '0';
    NumEdit := THNumEdit(GetControl('MTREGPER2' + IntToStr(i)));
    if NumEdit <> nil then MtRegCotPer2 := FloatToStr(NumEdit.Value);
    Motif := '';
    Edit := THEdit(GetControl('MOTIFABS' + IntToStr(i)));
    if Edit <> nil then Motif := Edit.text;
    NbHeures := '0';
    Edit := THEdit(GetControl('NBHEURREEL' + IntToStr(i)));
    if Edit <> nil then if Edit.text <> '' then NbHeures := Edit.text;
    NbHeureCompl := '0';
    Edit := THEdit(GetControl('NBHEURCOMPL' + IntToStr(i)));
    if Edit <> nil then if Edit.text <> '' then NbHeureCompl := Edit.text;
    NumEdit := THNumEdit(GetControl('MTSALRET' + IntToStr(i)));
    if NumEdit <> nil then MtSalRet := FloatToStr(NumEdit.Value);
    Check := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(i)));
    if Check <> nil then
      if Check <> nil then if Check.checked = True then SalRetEuro := 'X' else SalRetEuro := '-';
    {PT6 Utilisation d'une tob pour maj des champs}
    T := Tob.create('ATTSALAIRES', Tob_Mere, -1);
    if Ordre > 1 then T.SelectDB('"' + IntToStr(Ordre) + '";"' + IntToStr(i) + '"', nil); //PT11-2
    if T = nil then exit;
    T.PutValue('PAL_SALARIE', Matricule);
    T.PutValue('PAL_TYPEATTEST', 'MAL');
    T.PutValue('PAL_ORDRE', ordre);
    T.PutValue('PAL_MOIS', i);
    T.PutValue('PAL_DATEDEBUT', StrToDate(DateDebPer));
    T.PutValue('PAL_DATEFIN', StrToDate(DateFinPer));
    T.PutValue('PAL_SALAIRE', StrToFloat(StSal));
    T.PutValue('PAL_MONNAIESAL', SalEuro);
    T.PutValue('PAL_REGCOTPER1', MtRegCotPer1);
    T.PutValue('PAL_REGCOTPER2', MtRegCotPer2);
    T.PutValue('PAL_MOTIF', Motif);
    T.PutValue('PAL_NBHEURES', NbHeures);
    T.PutValue('PAL_NBHEURESCOMPL', NbHeureCompl);
    T.PutValue('PAL_SALAIRERET', MtSalRet);
    T.PutValue('PAL_MONNAIESALRET', SalRetEuro);
    {PT6 Mise en commentaire puis supprimé }
  end; //End For
  if Tob_Mere <> nil then
  begin
    if Tob_Mere.detail.count > 0 then
      Tob_Mere.InsertOrUpdateDB;
    Tob_Mere.free;
  end;
  {FIN PT6}
end;

procedure TOF_PGMALADIE.InitTaille(ControlParent: TWinControl);
var
  i: integer;
  ChampTrouve: string;
  ChampControl: TControl;
  Edit: THEdit;
  NumEdit: THNumEdit;
begin
    for i := 0 to ControlParent.ControlCount - 1 do
  begin
    ChampControl := ControlParent.Controls[i];
    ChampTrouve := AnsiUpperCase(ChampControl.Name);
    if ChampControl is THEdit then
    begin
      Edit := THEdit(ChampControl);
      Edit.Height := 17;
    end;
    if ChampControl is THNumEdit then
    begin
      NumEdit := THNumEdit(ChampControl);
      NumEdit.Height := 17;
    end;
  end;
end;

procedure TOF_PGMALADIE.OnLoad;
var
  PDFMaladie: TPDFPreview;
begin
  Loaded := False;
  PDFMaladie := TPDFPreview(GetControl('PDFMALADIE'));
  if PDFMaladie <> nil then InitTaille(PDFMaladie);
  SetControlEnabled('DATENAISSANCE', GetControlText('CKMOTIFMAL') <> 'X'); //PT11-1
  if GetControlText('CKMOTIFMAL') = 'X' then //SetControlText('DATENAISSANCE', '');
    SetControlText('DATENAISSANCE', DateToStr (idate1900)); //PT20
end;


procedure TOF_PGMALADIE.AffectDateArret(Sender: TObject);
begin
  {Edit:=THEdit(GetControl('DATEDERTRAV'));  PT2-3 mise en commentaire
  ArretEdit:=THEdit(GetControl('DATEARRET'));
  if (ArretEdit<>nil) and (Edit<>nil) then
     ArretEdit.Text:=Edit.Text;}
  if IsValidDate(GetControlText('DATEDERTRAV')) then //PT2-3 Modif code
    SetControlText('DATEARRET', DateToStr(StrToDate(GetControlText('DATEDERTRAV')) + 1));
end;

{PT6 n'est plus utilisé
function TOF_PGMALADIE.RemplaceVirguleParPoint(St: String): String;
var
i : Integer;
begin
For i:=1 to Length(St) do
   If St[i]=',' then St[i]:='.';
result := st;
end;}

procedure TOF_PGMALADIE.OnNew;
var
  Check: TcheckBox;
  Edit: THEdit;
  NumEdit: THNumEdit;
  Rep, i: integer;
begin
  Rep := PGIAsk('Voulez-vous saisir une nouvelle attestation?', Ecran.Caption);
  if rep = mrNo then exit;
  if rep = mrCancel then exit;
  if rep = mryes then
  begin
    Check := TcheckBox(GetControl('ETTEMPOTRAVAIL'));
    if check <> nil then Check.Checked := False;
    Edit := THEdit(GetControl('DATEDERTRAV'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DATEARRET'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DATEREPTRAV'));
    if edit <> nil then Edit.text := '';
    Check := TcheckBox(GetControl('CKMOTIFMAL'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CKMOTIFMAT'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CKMOTIFPAT'));
    if check <> nil then Check.Checked := False; //PT2-1
    Edit := THEdit(GetControl('SITUATIONARRET'));
    if edit <> nil then Edit.text := '';
    Check := TcheckBox(GetControl('CKNONREPRIS'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CKMOTIFPER'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CKMOTIFMED'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CKCASGEN'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CKCASPART'));
    if check <> nil then Check.Checked := False;
    NumEdit := THNumEdit(GetControl('MTCOT200'));
    if NumEdit <> nil then NumEdit.Value := 0;
    NumEdit := THNumEdit(GetControl('MTCOT800'));
    if NumEdit <> nil then NumEdit.Value := 0;
    Check := TcheckBox(GetControl('CK200'));
    if check <> nil then Check.Checked := False;
    Check := TcheckBox(GetControl('CK800'));
    if check <> nil then Check.Checked := False;
    Edit := THEdit(GetControl('DATEDEBPER'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DATEFINPER'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DATEFAIT'));
    if edit <> nil then Edit.text := DateToStr(Date);
    Edit := THEdit(GetControl('NOMSIGNE'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DATEDEBSUBR'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DATEFINSUBR'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('DENOMINATION'));
    if edit <> nil then Edit.text := '';
    Edit := THEdit(GetControl('INTITULE'));
    if edit <> nil then Edit.text := '';
    Check := TcheckBox(GetControl('PAIEEURO'));
    if check <> nil then Check.Checked := False;

    for i := 1 to 12 do
    begin //B For
      Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
      if Edit <> nil then Edit.Text := '';
      Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
      if Edit <> nil then Edit.Text := '';
      NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := 0;
      Check := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := False;
        Check.Visible := False;
      end;
      NumEdit := THNumEdit(GetControl('MTREGPER1' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := 0;
      NumEdit := THNumEdit(GetControl('MTREGPER2' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := 0;
      Edit := THEdit(GetControl('MOTIFABS' + IntToStr(i)));
      if Edit <> nil then Edit.text := '';
      Edit := THEdit(GetControl('NBHEURREEL' + IntToStr(i)));
      if Edit <> nil then Edit.text := '';
      Edit := THEdit(GetControl('NBHEURCOMPL' + IntToStr(i)));
      if Edit <> nil then Edit.text := '';
      NumEdit := THNumEdit(GetControl('MTSALRET' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := 0;
      Check := TCheckBox(GetControl('CKSALRETEURO' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := False;
        Check.Visible := False;
      end;
    end;
  end; //Mr yes
end;

procedure TOF_PGMALADIE.AfficheCheck(Sender: TObject);
var
  NumEdit: THNumEdit;
  check: TCheckBox;
  NomCheck, StI: string;
begin
  NumEdit := THNumEdit(Sender);
  if NumEdit = nil then exit;
  NomCheck := 'CKSALEURO';
  if Pos('MTSALRET', NumEdit.Name) = 1 then NomCheck := 'CKSALRETEURO' else NomCheck := 'CKSALEURO';
  StI := Copy(NumEdit.name, Length(NumEdit.Name), 1);
  check := TCheckBox(GetControl(NomCheck + StI));
  if Check = nil then exit;
  if NumEdit.value <> 0 then
    check.Visible := True
  else
  begin
    check.Checked := False;
    check.Visible := False;
  end;
end;

{ DEB PT12 }
procedure TOF_PGMALADIE.AffichDeclarant(Sender: TObject);
Var St : String;
begin
if GetControlText('DECLARANT')='' then exit;
St := RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False);
SetControlText('NOMSIGNE'        ,RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetControlText('LIEUFAIT'        ,RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
St := RechDom('PGDECLARANTQUAL'  ,GetControlText('DECLARANT'),False);
if St = 'AUT' then SetControlText('QUALITESIGNE' ,RechDom('PGDECLARANTAUTRE' ,GetControlText('DECLARANT'),False))
else               SetControlText('QUALITESIGNE' ,RechDom('PGQUALDECLARANT2' ,St,False));
end;
{ FIN PT12 }

{ DEB PT14 }
function TOF_PGMALADIE.SaisieNonValide: Boolean;
begin
  Result := False;
  if (GetControlText('CKMOTIFMAL') = '-')
  AND (GetControlText('CKMOTIFMAT') = '-')
  AND (GetControlText('CKMOTIFPAT') = '-') then
    Begin
    result := True;
    PgiBox('Vous devez renseigner le type d''attestation.',Ecran.Caption);
    End;
end;
{ FIN PT14 }

{ DEB PT16 13/09/2005 }
procedure TOF_PGMALADIE.Ck200Click(Sender: TObject);
Var DateDeb, DateFin : TDateTime;
begin
IF GetControlText('CK200') = 'X' then
  Begin
  SetControlText('MTCOT200','');
  SetControlEnabled('MTCOT200',False);
  SetControlText('CKMT200','-');
  SetControlEnabled('CKMT200',False);
  if IsValidDate(GetControlText('DATEARRET')) then
    Begin
    DateDeb := StrToDate(GetControlText('DATEARRET'));
    DateFin := DebutdeMois(DateDeb)-1;
    //Affectation "Montant de la cotisation"
    if (GetControlText('CKCASGEN')= 'X' ) then DateDeb := DebutDeMois(PlusMois(DebutDeMois(DateDeb)-1, -2))
    else
    if (GetControlText('CKCASPART')= 'X' ) then DateDeb := DebutDeMois(PlusMois(DebutDeMois(DateDeb)-1, -11));
    SetControlText('DATEDEBPER',DateToStr(DateDeb));
    SetControlText('DATEFINPER',DateToStr(DateFin));
    End;
  End
else
  Begin
  SetControlEnabled('MTCOT200',True);
  SetControlEnabled('CKMT200',True);
  AffectSalaire(nil);
  End;
end;
{ FIN PT16 13/09/2005 }

{ DEB PT16 13/09/2005 }
procedure TOF_PGMALADIE.Ck800Click(Sender: TObject);
Var DateDeb, DateFin : TDateTime;
begin
IF GetControlText('CK800') = 'X' then
  Begin
  SetControlText('MTCOT800','');
  SetControlEnabled('MTCOT800',False);
  SetControlText('CKMT800','-');
  SetControlEnabled('CKMT800',False);
  if IsValidDate(GetControlText('DATEARRET')) then
    Begin
    DateDeb := StrToDate(GetControlText('DATEARRET'));
    DateFin := DebutdeMois(DateDeb)-1;
    DateDeb := DebutDeMois(PlusMois(DebutDeMois(DateDeb)-1, -11));
    SetControlText('DATEDEBPER',DateToStr(DateDeb));
    SetControlText('DATEFINPER',DateToStr(DateFin));
    End;
  End
else
  Begin
  SetControlEnabled('MTCOT800',True);
  SetControlEnabled('CKMT800',True);
  AffectSalaire(nil);
  End;
end;
{ FIN PT16 13/09/2005 }

//DEB PT19
procedure TOF_PGMALADIE.MotifAbsClick(Sender: TObject);
var
  ThEdit1: THEdit;
  ThNumEdit2:THNumEdit;
  StI: string;
  St,Salarie:String;
  Q:TQuery;
  DateDeb,DateFin:TDateTime;
  Somme:Double;
begin
  ThEdit1 := THEdit(Sender);
  if ThEdit1 = nil then exit;
  StI := Copy(ThEdit1.name, Length(ThEdit1.Name), 1);
  if (ThEdit1.Text <> '') then
  begin
    Salarie := GetControlText('PSASALARIE');
    if (GetControlText('DATEDEBPER' + stI) <> '') and (GetControlText('DATEFINPER' + stI) <> '') then
    begin
      DateDeb := StrToDateTime(GetControlText('DATEDEBPER' + stI));
      DateFin := StrToDateTime(GetControlText('DATEFINPER' + stI));
      St := ' SELECT SUM (PHC_MONTANT) AS TOTAL'+
        ' FROM HISTOCUMSAL WHERE'+
        ' PHC_SALARIE="' + Salarie + '" AND'+
        ' PHC_CUMULPAIE="20" AND'+
        ' PHC_DATEFIN<="'+UsDateTime(DateFin)+'" AND'+
        ' PHC_DATEDEBUT>="'+UsDateTime(DateDeb)+'"';
      Q := OpenSQL(St,True);
      if not Q.Eof then
        SetControlText('NBHEURREEL' + stI,FloatToStr(Q.FindField('TOTAL').AsFloat));
      Ferme(Q);
    end;

    St := 'SELECT PSA_HORAIREMOIS,PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,' +
      'PSA_SALAIREMOIS5 FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"';
    Q := OpenSQL(St,True);
    if not Q.Eof then
    begin
      SetControlText('NBHEURCOMPL' + stI,FloatToStr(Q.FindField('PSA_HORAIREMOIS').AsFloat));
      Somme := Q.FindField('PSA_SALAIREMOIS1').AsFloat + Q.FindField('PSA_SALAIREMOIS2').AsFloat + Q.FindField('PSA_SALAIREMOIS3').AsFloat + Q.FindField('PSA_SALAIREMOIS4').AsFloat + Q.FindField('PSA_SALAIREMOIS5').AsFloat;
      SetControlText('MTSALRET' + stI,FloatToStr(Somme));
    end;
    Ferme(Q);
  end;
end;
//FIN PT19

initialization
  registerclasses([TOF_PGMALADIE]);
end.

