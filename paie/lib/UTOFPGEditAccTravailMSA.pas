{***********UNITE*************************************************
Auteur  ...... : PAIE PGI / RMA
Créé le ...... : 19/04/2007
Modifié le ... :   /  /
Description .. : gestion attestation Accident du Travail pour la MSA
Mots clefs ... : PAIE;ATTESTATIONS
*****************************************************************
}

unit UTOFPGEditAccTravailMSA;

interface
uses StdCtrls, Controls, Classes,sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  UtileAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat, ParamSoc,
  HPdfviewer, Vierge,
  Utob,
  PgOutilsTreso, Commun;

type
  TOF_PGACCTRAVAIL_MSA = class(TOF)
  Private
    Loaded, Ctrl, subrogation : Boolean;
    Matricule,GblEtab : string;
    Mode,DerjourTravSaisie: string;

    Procedure DateElipsisclick(Sender: TObject);
    Procedure InitSalarie;
    Procedure InitAttes;
    Procedure AffectSalaire;
    Procedure AffichDeclarant (Sender: TObject);
    procedure GestionSubrogation(Sender: TObject);
    Procedure EditExit(Sender: TObject);
    Procedure ChgtDate (Sender : TObject);
    Procedure CheckExit(Sender: TObject);
    Procedure Controle_periode(var ValPer1,Valper2 : String);
    Procedure Imprimer(Sender: TObject);
    Procedure UpdateTable(Sender: TObject);
    Procedure InitPages(Var Complet : Boolean);
    Function  FImpr(St : string): string;
  Public
    Procedure OnArgument(Arguments: string); override;
    Procedure OnLoad; override;
    Procedure OnNew; override;
  End;

implementation

uses PgEditOutils, EntPaie, P5Util ;

procedure TOF_PGACCTRAVAIL_MSA.OnArgument(Arguments: string);
var
  Edit     : THEdit;
  i        : integer;
  Check    : TCheckBox;
  St       : String;
  Page     : TPageControl;
  QAttes   : TQuery;
Begin
  inherited;
  Matricule := ReadTokenSt(Arguments);
  Mode := ReadTokenSt(Arguments);

  Page := TPageControl(GetControl('PAGES'));
  If Page = nil Then Exit;
  Page.ActivePageIndex := 0;

  If Trim(Matricule) <> '' then InitSalarie
                           Else Exit;

  TFVierge(Ecran).BImprimer.OnClick := Imprimer;
  TFVierge(Ecran).BValider.OnClick  := UpdateTable;

  //Evenement Volet 1
  Check := TCheckBox(GetControl('CKVICTOUI'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKVICTNON'));
  if Check <> nil then Check.OnClick := CheckExit;

  Edit := THEdit(GetControl('DATEACCTRAV'));
  If Edit <> nil then
  Begin
     Edit.OnElipsisClick := DateElipsisclick;
     Edit.OnChange := ChgtDate;
  End;
  Edit := THEdit(GetControl('DATEDERTRAV'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
    Edit.OnChange := ChgtDate;
  end;
  Edit := THEdit(GetControl('DATEREPTRAV'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
    Edit.OnChange := ChgtDate;
  end;
  Edit := THEdit(GetControl('HEUREACCTRAV'));
  If Edit <> nil then Edit.OnChange := ChgtDate;

  Check := TCheckBox(GetControl('CKMOTIFACCTRAV'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKMOTIFMALPROF'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKNONREPRIS'));
  if Check <> nil then Check.OnClick := CheckExit;

  For i := 1 to 4 do
  Begin
    Edit := THEdit(GetControl('DATEECHEANCE' + IntToStr(i)));
    If Edit <> nil Then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnDblClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnDblClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
  End;

  //Evenement Volet 2
  For i := 1 to 10 do
  Begin
    Edit := THEdit(GetControl('DATEVER' + IntToStr(i)));
    If Edit <> nil Then Edit.OnElipsisClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEDEBVER' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('DATEFINVER' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
  End;

  Edit := THEdit(GetControl('DATDEBPART'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATFINPART'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;

  Edit := THEdit(GetControl('DATDEBACTIV'));
  If Edit <> Nil then
  Begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnChange := ChgtDate;
  End;
  Edit := THEdit(GetControl('DATCTAPP'));
  If Edit <> Nil then
  Begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnChange := ChgtDate;
  End;
  Edit := THEdit(GetControl('DATEMBOCCAS'));
  If Edit <> Nil then
  Begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnChange := ChgtDate;
  End;
  Edit := THEdit(GetControl('DATDEBOCCAS'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATFINOCCAS'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;

  For i := 1 to 3 do
  Begin
    Edit := THEdit(GetControl('DATEDEBARRET' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('DATEFINARRET' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
  End;

  Check := TCheckBox(GetControl('MAINTIENOUI'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('MAINTIENNON'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('MAINTIENINT'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('MAINTIENPART'));
  if Check <> nil then Check.OnClick := CheckExit;

  Edit := THEdit(GetControl('DATEDEBSUBR'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATEFINSUBR'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DENOMINATION'));
  If Edit <> nil then Edit.OnChange := GestionSubrogation;
  Edit := THEdit(GetControl('DATEFAIT'));
  If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;

  St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GblEtab+'%") '+
      ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  ' ;
  SetControlProperty('DECLARANT', 'Plus' ,St );

  Edit := THEdit(GetControl('DECLARANT'));
  If Edit <> nil then Edit.OnExit := AffichDeclarant;

  If Mode = 'MODIFICATION' Then
  Begin
      If (ExisteSQL('SELECT PAS_SALARIE FROM ATTESTATIONS WHERE PAS_SALARIE="' + Matricule + '"' +
         'And PAS_TYPEATTEST = "MSA" And PAS_ASSEDICCAISSE = "ACT"') = TRUE) Then
          // Je me sert de PAS_ASSEDICCAISSE pour MAL=Attestation Maladie et ACT = Attestation AT

          InitAttes
      Else
          Mode := 'CREATION';
  End;
  If Mode = 'CREATION' Then
  Begin
      QAttes:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                    'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GblEtab+'%") '+
                    'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  '+
                    'ORDER BY PDA_ETABLISSEMENT DESC',True);
      If Not QAttes.eof then
      Begin
           SetControlText('DECLARANT' ,QAttes.FindField('PDA_DECLARANTATTES').AsString);
           AffichDeclarant(nil);
           SetControlText('DATEFAIT', DateToStr(Date));
      End;
      Ferme(QAttes);
  End;
  SetPlusBanqueCP (GetControl ('DENOMINATION'));
End;

Procedure TOF_PGACCTRAVAIL_MSA.InitSalarie;
var
  WDate  : Word;
  QSal   : TQuery;
  Defaut : THLabel;
Begin
  subrogation := False;

  Defaut := THLabel(GetControl('SOLIBELLE'));
  if Defaut <> nil then Defaut.Caption := GetParamSoc('SO_LIBELLE');

  QSal := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_NOMJF,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,' +
       'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,PSA_DATEANCIENNETE,' +
       'PSA_DATENAISSANCE,PSA_SEXE,PSA_NATIONALITE,PSA_COMMUNENAISS,PSA_DATEENTREE,PSA_QUALIFICATION,' +
       'ET_ETABLISSEMENT,ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,' +
       'ET_TELEPHONE,ET_SIRET,ETB_ETABLISSEMENT,ETB_NUMMSA,ETB_ACTIVITE,ETB_SUBROGATION '+
       'FROM SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
       'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
       'WHERE PSA_SALARIE="' + Matricule + '"', TRUE);

  If not QSal.EOF Then
  Begin
    GblEtab := QSal.FindField('ET_ETABLISSEMENT').asstring;
    Defaut  := THLabel(GetControl('ETLIBELLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_LIBELLE').asstring;
    Defaut := THLabel(GetControl('ETADRESSE123'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE1').asstring + ' ' + QSal.FindField('ET_ADRESSE2').asstring + ' ' + QSal.FindField('ET_ADRESSE3').asstring;
    Defaut := THLabel(GetControl('ETCP'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring, 'STR', 5);
    Defaut := THLabel(GetControl('ETVILLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_VILLE').asstring;
    Defaut := THLabel(GetControl('ETTELEPHONE'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_TELEPHONE').asstring, 'STR', 2);
    Defaut := THLabel(GetControl('ETSIRET'));
    if defaut <> nil then
       If Trim(QSal.FindField('ETB_NUMMSA').asstring) <> '' Then Defaut.caption := QSal.FindField('ETB_NUMMSA').asstring
          Else Defaut.caption := Copy(QSal.FindField('ET_SIRET').asstring,1,9);

    SetControlText('ETBACTIVITE',QSal.FindField('ETB_ACTIVITE').asstring);
    Defaut := THLabel(GetControl('PSASECU1'));
    if (defaut <> nil) then defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 1, 13), 'STR', 4);
    Defaut := THLabel(GetControl('PSASECU2'));
    if (defaut <> nil) then Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 14, 2), 'STR', 3);
    Defaut := THLabel(GetControl('PSADATENAISS'));
    if defaut <> nil then Defaut.caption := FImpr(QSal.FindField('PSA_DATENAISSANCE').asstring);
    Defaut := THLabel(GetControl('PSALIEUNAISS'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_COMMUNENAISS').asstring;

    If (QSal.FindField('PSA_SEXE').asstring = 'F') And (Trim(QSal.FindField('PSA_NOMJF').asstring) <> '') then
    Begin
       SetControlText('PSALIBELLE', QSal.FindField('PSA_NOMJF').asstring);
       SetControlText('PSALIBELLE1', QSal.FindField('PSA_LIBELLE').asstring);
    End
    Else
       SetControlText('PSALIBELLE', QSal.FindField('PSA_LIBELLE').asstring);

    SetControlText('PSAPRENOM', QSal.FindField('PSA_PRENOM').asstring);

    Defaut := THLabel(GetControl('PSAADRESSE123'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_ADRESSE1').asstring + '  ' + QSal.FindField('PSA_ADRESSE2').asstring + '  ' + QSal.FindField('PSA_ADRESSE3').asstring;
    Defaut := THLabel(GetControl('PSACP'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring, 'STR', 5);
    Defaut := THLabel(GetControl('PSAVILLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_VILLE').asstring;
    Defaut := THLabel(GetControl('PSADATEEMB'));
    if defaut <> nil then Defaut.caption := FImpr(QSal.FindField('PSA_DATEENTREE').asstring);
    Defaut := THLabel(GetControl('PSAEMPLOI'));
    if defaut <> nil then Defaut.caption := RechDom('PGLIBEMPLOI', QSal.FindField('PSA_LIBELLEEMPLOI').asstring, False);
    Defaut := THLabel(GetControl('PSAQUALIF'));
    if defaut <> nil then Defaut.caption := RechDom('PGLIBQUALIFICATION', QSal.FindField('PSA_QUALIFICATION').asstring, False);
    Defaut := THLabel(GetControl('PSAANC'));
    If defaut <> nil then
       If (QSal.FindField('PSA_DATEANCIENNETE').AsDateTime) > IDate1900 then
       Begin
         WDate := AncienneteAnnee(QSal.FindField('PSA_DATEANCIENNETE').AsDateTime, date);
         If WDate > 1 Then Defaut.caption := IntToStr(WDate) + ' ans'
         Else Defaut.caption := IntToStr(WDate) + ' an';
       End;

    If QSal.FindField('PSA_SEXE').asstring = 'M' Then
    Begin
       SetControlChecked('CKSEXM',True);
       SetControlChecked('CKSEXF',False);
    End
    Else
    Begin
       SetControlChecked('CKSEXM',False);
       SetControlChecked('CKSEXF',True);
    End;

    If Trim(QSal.FindField('PSA_NATIONALITE').asstring) <> '' Then
    Begin
      If (QSal.FindField('PSA_NATIONALITE').asstring) = 'FRA' Then
      Begin
         SetControlChecked('CKNATFRANC',True);
         SetControlChecked('CKNATCEE',False);
         SetControlChecked('CKNATAUTRE',False);
      End
      Else
      Begin
         If (RechDom('PGCEE', (QSal.FindField('PSA_NATIONALITE').asstring), False) <> '') Then
         Begin
            SetControlChecked('CKNATFRANC',False);
            SetControlChecked('CKNATCEE',True);
            SetControlChecked('CKNATAUTRE',False);
         End
         Else
         Begin
            SetControlChecked('CKNATFRANC',False);
            SetControlChecked('CKNATCEE',False);
            SetControlChecked('CKNATAUTRE',True);
         End;
      End;
    End;
    subrogation := QSal.FindField('ETB_SUBROGATION').Asstring = 'X';
    If (subrogation = True) Then
    Begin
       SetControlChecked('MAINTIENOUI',True);
       SetControlChecked('MAINTIENNON',False);
    End;
  End;
  Ferme(QSal);
End;

Procedure TOF_PGACCTRAVAIL_MSA.AffectSalaire;
Var
  QRech, QPlus : TQuery;
  Edit , Zone  : THedit;
  DerJour : TDateTime;
  j : integer;
  Pinit : boolean;
  UsDateDeb, UsDateFin: string;
  Mt : Double;
begin
  if Loaded = True then exit;
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if (Edit <> nil) then
  Begin
     If (IsValidDate(Edit.text)) And (Edit.text <> DerjourTravSaisie) Then
     Begin
        Pinit := False;
        InitPages(Pinit);
        DerJour := StrToDate(Edit.text);
        UsDateDeb := UsDateTime(DebutDeMois(PlusMois(DerJour, -1)));
        UsDateFin := UsDateTime(DerJour);
        DerjourTravSaisie := Edit.text;

        QRech := OpenSql('SELECT DISTINCT PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,' +
        'SUM(PPU_CBRUTFISCAL) AS PPU_CBRUTFISCAL, SUM(PPU_OCBRUTFISCAL) AS PPU_OCBRUTFISCAL,' +
        'SUM(PPU_CCOUTSALARIE) AS PPU_CCOUTSALARIE, SUM(PPU_OCCOUTSALARIE) AS PPU_OCCOUTSALARIE,' +
        'SUM(PPU_CHEURESTRAV) AS PPU_CHEURESTRAV ' +
        'FROM PAIEENCOURS ' +
        'WHERE PPU_SALARIE="' + Matricule + '" ' +
        'AND PPU_DATEDEBUT>="' + UsDateDeb + '" AND PPU_DATEFIN<"' + UsDateFin + '" ' +
        'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE ORDER BY PPU_DATEDEBUT DESC,PPU_DATEFIN DESC', TRUE);

        j := 0;
        QRech.First;
        While not QRech.Eof do
        Begin
           j := j + 1;
           Zone := THEdit(GetControl('DATEECHEANCE' + IntToStr(j)));
           If Zone <> nil Then Zone.text := QRech.FindField('PPU_PAYELE').asstring;
           Zone := THEdit(GetControl('DATEDEBPER' + IntToStr(j)));
           If Zone <> nil Then Zone.text := QRech.FindField('PPU_DATEDEBUT').asstring;
           Zone := THEdit(GetControl('DATEFINPER' + IntToStr(j)));
           If Zone <> nil then Zone.text := QRech.FindField('PPU_DATEFIN').asstring;
           Zone := THEdit(GetControl('HTRAV' + IntToStr(j)));
           If Zone <> nil Then Zone.text := QRech.FindField('PPU_CHEURESTRAV').asstring;
           If (VH_Paie.PGDateBasculEuro <= DebutDeMois(PlusMois(DerJour, j - 1))) Then
                 Mt := QRech.FindField('PPU_CBRUTFISCAL').asFloat
             Else
                 Mt := QRech.FindField('PPU_OCBRUTFISCAL').asFloat;

           SetControlText('MTSAL' + IntToStr(j), FloatToStr(Mt));
           SetFocusControl('MTSAL' + IntToStr(j)); //Evite de perdre la saisie du THNumEdit en cours si Mouette Verte direct

           QPlus := OpenSql('SELECT SUM(PHB_MTSALARIAL) AS CRDS,' +
                    'SUM(PHB_OMTSALARIAL) AS OCRDS FROM HISTOBULLETIN ' +
                   'LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE = PCT_RUBRIQUE AND PHB_NATURERUB = PCT_NATURERUB '+
                   'WHERE PHB_NATURERUB="COT" AND PCT_ETUDEDROIT = "X" '+
                   'AND PHB_SALARIE="' + Matricule + '" ' +
                   'AND PHB_DATEDEBUT="' + USDateTime(StrToDate(GetControlText('DATEDEBPER' + IntToStr(j)))) + '" '+
                   'AND PHB_DATEFIN="' + USDateTime(StrToDate(GetControlText('DATEFINPER' + IntToStr(j)))) + '" ',True);

           If Not QPlus.eof Then
           Begin
              If (VH_Paie.PGDateBasculEuro <= DebutDeMois(PlusMois(DerJour, j - 1))) Then
                 Mt := QRech.FindField('PPU_CCOUTSALARIE').asFloat - QPlus.FindField('CRDS').asFloat
              Else
                 Mt := QRech.FindField('PPU_OCCOUTSALARIE').asFloat - QPlus.FindField('OCRDS').asFloat;

              SetControlText('PARTSAL' + IntToStr(j), FloatToStr(Mt));
              SetFocusControl('PARTSAL' + IntToStr(j)); //Evite de perdre la saisie du THNumEdit en cours si Mouette Verte direct
           End;
           Ferme(QPlus);
           If j = 4 Then Break;
           QRech.Next;
        End;
        Ferme(QRech);
        SetFocusControl('HEUREACCTRAV');
     End;
  End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.AffichDeclarant(Sender: TObject);
Var
  St : String;
Begin
  If GetControlText('DECLARANT')='' Then exit;
  St := RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False);
  SetControlText('NOMSIGNE' ,RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
  SetControlText('LIEUFAIT' ,RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
  St := RechDom('PGDECLARANTQUAL'  ,GetControlText('DECLARANT'),False);
  If St = 'AUT' Then SetControlText('QUALITESIGNE' ,RechDom('PGDECLARANTAUTRE' ,GetControlText('DECLARANT'),False))
                Else SetControlText('QUALITESIGNE' ,RechDom('PGQUALDECLARANT2' ,St,False));
End;

Procedure TOF_PGACCTRAVAIL_MSA.GestionSubrogation(Sender: TObject);
var
  deno   : THedit;
  QRech  : Tquery;
  StPlus : String;
Begin
  //PORTAGECWAS
  SetControlEnabled('BANQUE', False);
  SetControlText('BANQUE', '');
  SetControlEnabled('GUICHET', False);
  SetControlText('GUICHET', '');
  SetControlEnabled('COMPTE', False);
  SetControlText('COMPTE', '');
  SetControlEnabled('CLE', False);
  SetControlText('CLE', '');

  deno := THedit(GetControl('DENOMINATION'));
  If deno <> nil then
    If deno.text <> '' then
    Begin
      StPlus:= PGBanqueCP (True);

      QRech := opensql('SELECT BQ_ETABBQ,BQ_NUMEROCOMPTE,BQ_CLERIB,BQ_GUICHET ' +
        'FROM BANQUECP WHERE BQ_GENERAL="' + deno.text + '"'+StPlus, TRUE);

      If not Qrech.eof Then //PORTAGECWAS
      Begin
         SetControlEnabled('BANQUE', True);
         SetControlText('BANQUE', QRech.FindField('BQ_ETABBQ').asstring);
         SetControlEnabled('GUICHET', True);
         SetControlText('GUICHET', QRech.FindField('BQ_GUICHET').asstring);
         SetControlEnabled('COMPTE', True);
         SetControlText('COMPTE', QRech.FindField('BQ_NUMEROCOMPTE').asstring);
         SetControlEnabled('CLE', True);
         SetControlText('CLE', QRech.FindField('BQ_CLERIB').asstring);
      End;
      Ferme(QRech);
    End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.EditExit(Sender: TObject);
var
  Valeur1, Valeur2 : String;
  i : integer;
Begin
  If THEdit(Sender).Name = 'DATEDERTRAV' Then
     If IsValidDate(GetControlText('DATEDERTRAV')) Then AffectSalaire;

  If THEdit(Sender).Name = 'DATEREPTRAV' Then
     If IsValidDate(GetControlText('DATEREPTRAV')) Then SetControlText('CKNONREPRIS','-');

  For i := 1 to 4 do
  Begin
    If (THEdit(Sender).Name = 'DATEDEBPER'+ IntToStr(i)) Or (THEdit(Sender).Name = 'DATEFINPER'+ IntToStr(i)) Then
    Begin
       Valeur1 := GetControlText('DATEDEBPER'+ IntToStr(i));
       Valeur2 := GetControlText('DATEFINPER'+ IntToStr(i));
       Controle_periode(Valeur1,Valeur2);
       If Ctrl = False Then SetFocusControl('DATEDEBPER'+ IntToStr(i));
    End;
  End;

  For i := 1 to 10 do
  Begin
    If (THEdit(Sender).Name = 'DATEDEBVER'+ IntToStr(i)) Or (THEdit(Sender).Name = 'DATEFINVER'+ IntToStr(i)) Then
    Begin
       Valeur1 := GetControlText('DATEDEBVER'+ IntToStr(i));
       Valeur2 := GetControlText('DATEFINVER'+ IntToStr(i));
       Controle_periode(Valeur1,Valeur2);
       If Ctrl = False Then SetFocusControl('DATEDEBVER'+ IntToStr(i));
    End;
  End;

  If (THEdit(Sender).Name = 'DATDEBPART') Or (THEdit(Sender).Name = 'DATFINPART') Then
  Begin
     Valeur1 := GetControlText('DATDEBPART');
     Valeur2 := GetControlText('DATFINPART');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DATDEBPART');
  End;

  If (THEdit(Sender).Name = 'DATDEBOCCAS') Or (THEdit(Sender).Name = 'DATFINOCCAS') Then
  Begin
     Valeur1 := GetControlText('DATDEBOCCAS');
     Valeur2 := GetControlText('DATFINOCCAS');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DATDEBOCCAS');
  End;

  For i := 1 to 3 do
  Begin
    If (THEdit(Sender).Name = 'DATEDEBARRET'+ IntToStr(i)) Or (THEdit(Sender).Name = 'DATEFINARRET'+ IntToStr(i)) Then
    Begin
       Valeur1 := GetControlText('DATEDEBARRET'+ IntToStr(i));
       Valeur2 := GetControlText('DATEFINARRET'+ IntToStr(i));
       Controle_periode(Valeur1,Valeur2);
       If Ctrl = False Then SetFocusControl('DATEDEBARRET'+ IntToStr(i));
    End;
  End;

  If (THEdit(Sender).Name = 'DATEDEBSUBR') Or (THEdit(Sender).Name = 'DATEFINSUBR') Then
  Begin
     Valeur1 := GetControlText('DATEDEBSUBR');
     Valeur2 := GetControlText('DATEFINSUBR');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DATEDEBSUBR');
  End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.ChgtDate(Sender: TObject);
var
  TXT : String;
Begin
  If THEdit(Sender).Name ='DATEACCTRAV' Then
  Begin
    TXT := FImpr(GetControlText('DATEACCTRAV'));
    SetControlText('DATEACCTRAV1', TXT);
  End;
  If THEdit(Sender).Name ='DATEDERTRAV' Then
  Begin
    TXT := FImpr(GetControlText('DATEDERTRAV'));
    SetControlText('DATEDERTRAV1', TXT);
  End;
  If THEdit(Sender).Name ='HEUREACCTRAV' Then
  Begin
    TXT := FImpr(GetControlText('HEUREACCTRAV'));
    SetControlText('HEUREACCTRAV1', TXT);
  End;
  If THEdit(Sender).Name ='DATEREPTRAV' Then
  Begin
    TXT := FImpr(GetControlText('DATEREPTRAV'));
    SetControlText('DATEREPTRAV1', TXT);
  End;
  If THEdit(Sender).Name ='DATEMBOCCAS' Then
  Begin
    TXT := FImpr(GetControlText('DATEMBOCCAS'));
    SetControlText('DATEMBOCCAS1', TXT);
  End;
  If THEdit(Sender).Name ='DATDEBACTIV' Then
  Begin
    TXT := FImpr(GetControlText('DATDEBACTIV'));
    SetControlText('DATDEBACTIV1', TXT);
  End;
  If THEdit(Sender).Name ='DATCTAPP' Then
  Begin
    TXT := FImpr(GetControlText('DATCTAPP'));
    SetControlText('DATCTAPP1', TXT);
  End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.CheckExit(Sender: TObject);
Begin
  If TCheckBox(Sender).Name = 'CKVICTOUI' Then
     IF GetControlText('CKVICTOUI') = 'X' then SetControlText('CKVICTNON','-');
  If TCheckBox(Sender).Name = 'CKVICTNON' Then
     IF GetControlText('CKVICTNON') = 'X' then SetControlText('CKVICTOUI','-');
  If TCheckBox(Sender).Name = 'CKMOTIFACCTRAV' Then
     IF GetControlText('CKMOTIFACCTRAV') = 'X' then SetControlText('CKMOTIFMALPROF','-');
  If TCheckBox(Sender).Name = 'CKMOTIFMALPROF' Then
     IF GetControlText('CKMOTIFMALPROF') = 'X' then SetControlText('CKMOTIFACCTRAV','-');
  If TCheckBox(Sender).Name = 'CKNONREPRIS' Then
     IF GetControlText('CKNONREPRIS') = 'X' then SetControlText('DATEREPTRAV','');
  If TCheckBox(Sender).Name = 'MAINTIENOUI' Then
     IF GetControlText('MAINTIENOUI') = 'X' then SetControlText('MAINTIENNON','-');
  If TCheckBox(Sender).Name = 'MAINTIENNON' Then
     IF GetControlText('MAINTIENNON') = 'X' then SetControlText('MAINTIENOUI','-');
  If TCheckBox(Sender).Name = 'MAINTIENINT' Then
     IF GetControlText('MAINTIENINT') = 'X' then SetControlText('MAINTIENPART','-');
  If TCheckBox(Sender).Name = 'MAINTIENPART' Then
     IF GetControlText('MAINTIENPART') = 'X' then SetControlText('MAINTIENINT','-');
End;

Procedure TOF_PGACCTRAVAIL_MSA.Controle_periode(var ValPer1,Valper2 : String);
Begin
  Ctrl := True;
  If (IsValidDate(ValPer1)) And (IsValidDate(ValPer2)) Then
  Begin
     If (StrToDate(ValPer1) > StrToDate(ValPer2)) And (ValPer2 <> '01/01/1900') Then
     Begin
        PGIBox ('La période est invalide', TFVierge(Ecran).caption);
        Ctrl := False;
     End;
  End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.UpdateTable(Sender: TObject);
var
  Check: TcheckBox;
  Edit : THEdit;
  NumEdit: THNumEdit;
  DerjourTrav, DateRepArret, MotifArret, NonRepris,DateAccTrav : string;
  DateDebPer, DateFinPer, LieuDecla, DateDecla, PersDecla, QualDecla : string;
  DateDebSubr, DateFinSubr, AutreVictime, HeureAcc : string;
  DATDEBPART,DATFINPART,DATDEBACTIV,DATCTAPP,DATEMBOCCAS,DATDEBOCCAS,DATFINOCCAS : String;
  Mntc001,Mntc002,Mntc003,Mntc004,Mntc005,Mntc006,Mntc007,Mntc008,Mntd001,Mntd002 : double;
  Maint, TMaint, Hoccas, Motif, DateDebArret, DateFinArret : string;
  CKintertrav, CKembacc,CompteBq,Htrav, Echeance, DateVer, DateDebVer, DateFinVer : String;
  MntSubro, StSal, Avant, Indem, Soumis, DedSup, PartSal, BrutVer, PartSalb : double;
  ordre, i : integer;
  QRech: TQuery;
  Tob_Mere, T: TOB;
begin
  ordre := 0;
  AutreVictime := '-';
  SetFocusControl('DATEFAIT'); //Evite de perdre la saisie du THNumEdit en cours si Mouette Verte direct
  Check := TcheckBox(GetControl('CKVICTOUI'));
  if check <> nil then if Check.Checked = True then AutreVictime := 'X';
  DerjourTrav := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if edit <> nil then if IsValidDate(Edit.text) then DerjourTrav := Edit.text;
  DateAccTrav := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEACCTRAV'));
  if edit <> nil then if IsValidDate(Edit.text) then DateAccTrav := Edit.text;
  DateRepArret := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEREPTRAV'));
  if edit <> nil then if IsValidDate(Edit.text) then DateRepArret := Edit.text;
  HeureAcc := '00:00';
  Edit := THEdit(GetControl('HEUREACCTRAV'));
  if edit <> nil then if Edit.text <> '' then HeureAcc := Edit.text;
  MotifArret := '';
  Check := TcheckBox(GetControl('CKMOTIFACCTRAV'));
  if check <> nil then if Check.Checked = True then MotifArret := 'ACC';
  Check := TcheckBox(GetControl('CKMOTIFMALPROF'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MPR';
  NonRepris := '-';
  Check := TcheckBox(GetControl('CKNONREPRIS'));
  if check <> nil then if Check.Checked = True then NonRepris := 'X';
  DATDEBPART := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATDEBPART'));
  if edit <> nil then if IsValidDate(Edit.text) then DATDEBPART := Edit.text;
  DATFINPART := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATFINPART'));
  if edit <> nil then if IsValidDate(Edit.text) then DATFINPART := Edit.text;
  DATDEBACTIV := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATDEBACTIV'));
  if edit <> nil then if IsValidDate(Edit.text) then DATDEBACTIV := Edit.text;
  DATCTAPP := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATCTAPP'));
  if edit <> nil then if IsValidDate(Edit.text) then DATCTAPP := Edit.text;
  DATEMBOCCAS := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEMBOCCAS'));
  if edit <> nil then if IsValidDate(Edit.text) then DATEMBOCCAS := Edit.text;
  DATDEBOCCAS := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATDEBOCCAS'));
  if edit <> nil then if IsValidDate(Edit.text) then DATDEBOCCAS := Edit.text;
  DATFINOCCAS := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATFINOCCAS'));
  if edit <> nil then if IsValidDate(Edit.text) then DATFINOCCAS := Edit.text;
  Mntc001 := 0;
  NumEdit := THNumEdit(GetControl('MNTC001'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc001 := NumEdit.value;
  Mntc002 := 0;
  NumEdit := THNumEdit(GetControl('MNTC002'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc002 := NumEdit.value;
  Mntc003 := 0;
  NumEdit := THNumEdit(GetControl('MNTC003'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc003 := NumEdit.value;
  Mntc004 := 0;
  NumEdit := THNumEdit(GetControl('MNTC004'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc004 := NumEdit.value;
  Mntc005 := 0;
  NumEdit := THNumEdit(GetControl('MNTC005'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc005 := NumEdit.value;
  Mntc006 := 0;
  NumEdit := THNumEdit(GetControl('MNTC006'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc006 := NumEdit.value;
  Mntc007 := 0;
  NumEdit := THNumEdit(GetControl('MNTC007'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc007 := NumEdit.value;
  Mntc008 := 0;
  NumEdit := THNumEdit(GetControl('MNTC008'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntc008 := NumEdit.value;
  Hoccas := '0';
  Edit := THEdit(GetControl('HOCCAS'));
  if edit <> nil then if Edit.text <> '' then Hoccas := Edit.text;
  Mntd001 := 0;
  NumEdit := THNumEdit(GetControl('MNTD001'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntd001 := NumEdit.value;
  Mntd002 := 0;
  NumEdit := THNumEdit(GetControl('MNTD002'));
  if NumEdit <> nil then if NumEdit.Value > 0 then Mntd002 := NumEdit.value;
  CKintertrav := '-';
  Check := TcheckBox(GetControl('CKINTERTRAV'));
  if check <> nil then if Check.Checked = True then CKintertrav := 'X';
  CKembacc := '-';
  Check := TcheckBox(GetControl('CKEMBACC'));
  if check <> nil then if Check.Checked = True then CKembacc := 'X';
  Maint := '-';
  Check := TcheckBox(GetControl('MAINTIENOUI'));
  if check <> nil then if Check.Checked = True then Maint := 'X';
  TMaint := '';
  Check := TcheckBox(GetControl('MAINTIENINT'));
  if check <> nil then if Check.Checked = True then TMaint := 'INT';
  Check := TcheckBox(GetControl('MAINTIENPART'));
  if check <> nil then if Check.Checked = True then TMaint := 'PAR';
  MntSubro := 0;
  NumEdit := THNumEdit(GetControl('SUBMT'));
  if NumEdit <> nil then if NumEdit.Value > 0 then MntSubro := NumEdit.value;
  DateDebSubr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDEBSUBR'));
  if edit <> nil then if IsValidDate(Edit.text) then DateDebSubr := Edit.text;
  DateFinSubr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFINSUBR'));
  if edit <> nil then if IsValidDate(Edit.text) then DateFinSubr := Edit.text;
  CompteBq := '';
  Edit := THEdit(GetControl('DENOMINATION'));
  if edit <> nil then if Edit.text <> '' then CompteBq := Edit.text;
  LieuDecla := '';
  Edit := THEdit(GetControl('LIEUFAIT'));
  if edit <> nil then if Edit.text <> '' then LieuDecla := Edit.text;
  DateDecla := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFAIT'));
  if edit <> nil then if IsValidDate(Edit.text) then DateDecla := Edit.text;
  PersDecla := '';
  Edit := THEdit(GetControl('NOMSIGNE'));
  if edit <> nil then if Edit.text <> '' then PersDecla := Edit.text;
  QualDecla := '';
  Edit := THEdit(GetControl('QUALITESIGNE'));
  if Edit <> nil then if Edit.Text <> '' then QualDecla := Edit.Text;

  if (subrogation = True) and
     ((DateDebSubr = DateToStr(Idate1900)) or
      (DateFinSubr = DateToStr(Idate1900)) or
      (CompteBq = '')) Then

     PgiBox('Des informations concernant la subrogation sont absentes!','Subrogation');

  QRech := OpenSql('SELECT PAS_ORDRE FROM ATTESTATIONS ' +
    'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="MSA" AND PAS_ASSEDICCAISSE = "ACT"', True);

  If not QRech.Eof then ordre := QRech.Fields[0].asinteger;
  Ferme(QRech);
  If ordre = 0 then
  Begin
    QRech := OpenSql('SELECT MAX(PAS_ORDRE) FROM ATTESTATIONS ', True);
    If not QRech.EOF then //PORTAGECWAS
      ordre := QRech.Fields[0].asinteger + 1
    Else
      ordre := 1;
    Ferme(QRech);
  End;

  Tob_Mere := Tob.create('ATTESTATIONS', nil, -1);
  T := Tob.create('ATTESTATIONS', Tob_Mere, -1);
  If Ordre > 1 then T.SelectDB('"' + Matricule + '";"MSA";"' + IntToStr(Ordre) + '"', nil);

  If T = nil then exit;
  T.PutValue('PAS_SALARIE', Matricule);
  T.PutValue('PAS_TYPEATTEST', 'MSA');
  T.PutValue('PAS_ORDRE', ordre);
  T.PutValue('PAS_ASSEDICCAISSE', 'ACT');                 //Champ détourné
  T.PutValue('PAS_AUTREVICTIME',AutreVictime);
  T.PutValue('PAS_DERNIERJOUR', StrToDate(DerjourTrav));
  T.PutValue('PAS_DATEARRET', StrToDate(DateAccTrav));
  T.PutValue('PAS_REPRISEARRET', StrToDate(DateRepArret));
  T.PutValue('PAS_ASSEDICNUM',HeureAcc);                  //Champ détourné
  T.PutValue('PAS_MOTIFARRET', MotifArret);
  T.PutValue('PAS_NONREPRIS', NonRepris);
  T.PutValue('PAS_PERIODEDEBUT',StrToDate(DATDEBPART));   //Champ détourné
  T.PutValue('PAS_PERIODEFIN',StrToDate(DATFINPART));     //Champ détourné
  T.PutValue('PAS_DATEPLANSOC',StrToDate(DATDEBACTIV));   //Champ détourné
  T.PutValue('PAS_DATEPREAVISD2',StrToDate(DATCTAPP));    //Champ détourné
  T.PutValue('PAS_DATEPREAVISF2',StrToDate(DATEMBOCCAS)); //Champ détourné
  T.PutValue('PAS_DATEPREAVISD3',StrToDate(DATDEBOCCAS)); //Champ détourné
  T.PutValue('PAS_DATEPREAVISF3',StrToDate(DATFINOCCAS)); //Champ détourné
  T.PutValue('PAS_INDPREAVIS',Mntc001);                   //Champ détourné
  T.PutValue('PAS_INDCPMONTANT',Mntc002);                 //Champ détourné
  T.PutValue('PAS_INDINHER',Mntc003);                     //Champ détourné
  T.PutValue('PAS_INDLEGM',Mntc004);                      //Champ détourné
  T.PutValue('PAS_INDCDDM',Mntc005);                      //Champ détourné
  T.PutValue('PAS_INDCNEM',Mntc006);                      //Champ détourné
  T.PutValue('PAS_INDAUTREM',Mntc007);                    //Champ détourné
  T.PutValue('PAS_INDFINMISSM',Mntc008);                  //Champ détourné
  T.PutValue('PAS_HORHEBENT',StrToFloat(Hoccas));         //Champ détourné
  T.PutValue('PAS_INDRETRAITEM',Mntd001);                 //Champ détourné
  T.PutValue('PAS_INDSINISTREM',Mntd002);                 //Champ détourné
  T.PutValue('PAS_RECLASS',CKintertrav);                  //Champ détourné
  T.PutValue('PAS_PLANSOC',CKembacc);                     //Champ détourné
  T.PutValue('PAS_MAINTIEN',Maint);
  T.PutValue('PAS_TYPMAINTIEN',TMaint);
  T.PutValue('PAS_MONTANT',MntSubro);
  T.PutValue('PAS_SUBDEBUT', StrToDate(DateDebSubr));
  T.PutValue('PAS_SUBFIN', StrToDate(DateFinSubr));
  T.PutValue('PAS_SUBCOMPTE', CompteBq);
  T.PutValue('PAS_DECLARLIEU', LieuDecla);
  T.PutValue('PAS_DECLARDATE', StrToDate(DateDecla));
  T.PutValue('PAS_DECLARPERS', PersDecla);
  T.PutValue('PAS_DECLARQUAL', QualDecla);
  T.InsertOrUpdateDB;
  If Tob_Mere <> nil Then Tob_Mere.free;

  //MAJ Tableau
  Tob_Mere := Tob.create('ATTSALAIRES', nil, -1);

  For i := 1 to 10 do
  Begin
    // Pavé A
    Echeance := DateToStr(Idate1900);
    DateDebPer := DateToStr(Idate1900);
    DateFinPer := DateToStr(Idate1900);
    Htrav := '0';
    StSal := 0;
    Avant := 0;
    Indem := 0;
    Soumis := 0;
    DedSup := 0;
    PartSal := 0;
    If i <= 4 Then
    Begin
       Edit := THEdit(GetControl('DATEECHEANCE' + IntToStr(i)));
       If Edit <> nil then if IsValidDate(Edit.text) then Echeance := Edit.text;
       Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
       if Edit <> nil then if IsValidDate(Edit.text) then DateDebPer := Edit.text;
       Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
       if Edit <> nil then if IsValidDate(Edit.text) then DateFinPer := Edit.text;
       Edit := THEdit(GetControl('HTRAV' + IntToStr(i)));
       if Edit <> nil then if Edit.Text <> '' then Htrav := Edit.text;
       NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
       if NumEdit <> nil then if NumEdit.Value > 0 Then StSal := NumEdit.value;
       NumEdit := THNumEdit(GetControl('AVANT' + IntToStr(i)));
       if NumEdit <> nil then if NumEdit.Value > 0 Then Avant := NumEdit.value;
       NumEdit := THNumEdit(GetControl('INDEM' + IntToStr(i)));
       if NumEdit <> nil then if NumEdit.Value > 0 Then Indem := NumEdit.value;
       NumEdit := THNumEdit(GetControl('SOUMIS' + IntToStr(i)));
       if NumEdit <> nil then if NumEdit.Value > 0 Then Soumis := NumEdit.value;
       NumEdit := THNumEdit(GetControl('DEDSUP' + IntToStr(i)));
       if NumEdit <> nil then if NumEdit.Value > 0 Then DedSup := NumEdit.value;
       NumEdit := THNumEdit(GetControl('PARTSAL' + IntToStr(i)));
       if NumEdit <> nil then if NumEdit.Value > 0 Then PartSal := NumEdit.value;
    End;
    // Pavé B
    DateVer := DateToStr(Idate1900);
    DateDebVer := DateToStr(Idate1900);
    DateFinVer := DateToStr(Idate1900);
    BrutVer := 0;
    PartSalb := 0;

    Edit := THEdit(GetControl('DATEVER' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DateVer := Edit.text;
    Edit := THEdit(GetControl('DATEDEBVER' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DateDebVer := Edit.text;
    Edit := THEdit(GetControl('DATEFINVER' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DateFinVer := Edit.text;
    NumEdit := THNumEdit(GetControl('BRUTVER' + IntToStr(i)));
    if NumEdit <> nil then if NumEdit.Value > 0 Then BrutVer := NumEdit.value;
    NumEdit := THNumEdit(GetControl('PARTSALB' + IntToStr(i)));
    if NumEdit <> nil then if NumEdit.Value > 0 Then PartSalb := NumEdit.value;

    // Pavé D
    Motif := '';
    DateDebArret := DateToStr(Idate1900);
    DateFinArret := DateToStr(Idate1900);
    If i <= 3 Then
    Begin
       Edit := THEdit(GetControl('MOTIFARRET' + IntToStr(i)));
       if Edit <> nil then if Edit.Text <> '' then Motif := Edit.text;
       Edit := THEdit(GetControl('DATEDEBARRET' + IntToStr(i)));
       if Edit <> nil then if IsValidDate(Edit.text) then DateDebArret := Edit.text;
       Edit := THEdit(GetControl('DATEFINARRET' + IntToStr(i)));
       if Edit <> nil then if IsValidDate(Edit.text) then DateFinArret := Edit.text;
    End;

    T := Tob.create('ATTSALAIRES', Tob_Mere, -1);
    if Ordre > 1 then T.SelectDB('"' + IntToStr(Ordre) + '";"' + IntToStr(i) + '"', nil);
    if T = nil then exit;
    T.PutValue('PAL_SALARIE', Matricule);
    T.PutValue('PAL_TYPEATTEST', 'MSA');
    T.PutValue('PAL_ORDRE', ordre);
    T.PutValue('PAL_MOIS', i);
    T.PutValue('PAL_ECHEANCE', StrToDate(Echeance));
    T.PutValue('PAL_DATEDEBUT', StrToDate(DateDebPer));
    T.PutValue('PAL_DATEFIN', StrToDate(DateFinPer));
    T.PutValue('PAL_NBHEURES',StrToFloat(Htrav));
    T.PutValue('PAL_SALAIRE', StSal);
    T.PutValue('PAL_AVANTNAT', Avant);
    T.PutValue('PAL_INDEMNITE', Indem);
    T.PutValue('PAL_SOUMIS', Soumis);
    T.PutValue('PAL_DEDSUP', DedSup);
    T.PutValue('PAL_PARTSAL4', PartSal);
    T.PutValue('PAL_VERSEMENT', StrToDate(DateVer));
    T.PutValue('PAL_DEBPERVER', StrToDate(DateDebVer));
    T.PutValue('PAL_FINPERVER', StrToDate(DateFinVer));
    T.PutValue('PAL_MONTANTVER', BrutVer);
    T.PutValue('PAL_PARTSAL12', PartSalb);
    T.PutValue('PAL_MOTIF', Motif);
    T.PutValue('PAL_DEBARRET', StrToDate(DateDebArret));
    T.PutValue('PAL_FINARRET', StrToDate(DateFinArret));
  End; //End For
  If Tob_Mere <> nil Then
  Begin
    If Tob_Mere.detail.count > 0 then
       Tob_Mere.InsertOrUpdateDB;
    Tob_Mere.free;
  End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.InitAttes;
var
  QAttes : TQuery;
  Edit   : THEdit;
  NumEdit: THNumEdit;
  Ordre, STrlu : string;
  i : integer;
Begin
  Loaded := True;
  QAttes := OpenSql('SELECT PAS_ORDRE,PAS_MOTIFARRET,PAS_DERNIERJOUR,PAS_REPRISEARRET,PAS_AUTREVICTIME,' +
     'PAS_ASSEDICNUM,PAS_DATEARRET,PAS_NONREPRIS,PAS_PLANSOC,PAS_RECLASS,PAS_DATEPREAVISD2,PAS_DATEPREAVISF2,' +
     'PAS_DATEPREAVISD3,PAS_DATEPREAVISF3,PAS_DATEPLANSOC,PAS_MONTANT,PAS_PERIODEDEBUT,PAS_PERIODEFIN,' +
     'PAS_INDCPMONTANT,PAS_INDINHER,PAS_INDLEGM,PAS_INDCDDM,PAS_INDCNEM,PAS_INDAUTREM,PAS_INDFINMISSM,' +
     'PAS_HORHEBENT,PAS_INDRETRAITEM,PAS_INDSINISTREM,PAS_MAINTIEN,' +
     'PAS_SUBDEBUT,PAS_SUBFIN,PAS_TYPMAINTIEN,PAS_INDPREAVIS,PAS_DECLARLIEU,PAS_DECLARDATE,' +
     'PAS_DECLARPERS,PAS_DECLARQUAL,PAS_SUBCOMPTE FROM ATTESTATIONS ' +
     'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="MSA" AND PAS_ASSEDICCAISSE="ACT"', True);

  If not QAttes.EOF Then //PORTAGECWAS
  Begin
     Ordre := QAttes.FindField('PAS_ORDRE').asstring;
     STrlu := QAttes.FindField('PAS_AUTREVICTIME').asstring;
     If STrlu = 'X' Then SetControlText('CKVICTOUI','X');
     If STrlu = '-' Then SetControlText('CKVICTNON','X');

     SetControlText('DATEACCTRAV1','');
     SetControlText('DATEDERTRAV1','');
     SetControlText('HEUREACCTRAV1','');
     SetControlText('DATEREPTRAV1','');
     SetControlText('DATEMBOCCAS1','');
     SetControlText('DATDEBACTIV1','');
     SetControlText('DATCTAPP1','');

     If (QAttes.FindField('PAS_DERNIERJOUR').AsDateTime <> IDate1900) then
     Begin
         SetControlText('DATEDERTRAV', DateToStr(QAttes.FindField('PAS_DERNIERJOUR').AsDateTime));
         DerjourTravSaisie := DateToStr(QAttes.FindField('PAS_DERNIERJOUR').AsDateTime);
     End;
     If (QAttes.FindField('PAS_DATEARRET').AsDateTime <> IDate1900) then
         SetControlText('DATEACCTRAV', DateToStr(QAttes.FindField('PAS_DATEARRET').AsDateTime));
     If (QAttes.FindField('PAS_REPRISEARRET').AsDateTime <> IDate1900) then
         SetControlText('DATEREPTRAV', DateToStr(QAttes.FindField('PAS_REPRISEARRET').AsDateTime));
     If (QAttes.FindField('PAS_ASSEDICNUM').asstring <> '00:00') then
         SetControlText('HEUREACCTRAV', QAttes.FindField('PAS_ASSEDICNUM').asstring);

     STrlu := QAttes.FindField('PAS_MOTIFARRET').asstring;
     If STrlu = 'ACC' Then SetControlText('CKMOTIFACCTRAV','X');
     If STrlu = 'MPR' Then SetControlText('CKMOTIFMALPROF','X');

     SetControlText('CKNONREPRIS',QAttes.FindField('PAS_NONREPRIS').asstring);

     If (QAttes.FindField('PAS_PERIODEDEBUT').AsDateTime <> IDate1900) then
         SetControlText('DATDEBPART', DateToStr(QAttes.FindField('PAS_PERIODEDEBUT').AsDateTime));
     If (QAttes.FindField('PAS_PERIODEFIN').AsDateTime <> IDate1900) then
         SetControlText('DATFINPART', DateToStr(QAttes.FindField('PAS_PERIODEFIN').AsDateTime));
     If (QAttes.FindField('PAS_DATEPLANSOC').AsDateTime <> IDate1900) then
         SetControlText('DATDEBACTIV', DateToStr(QAttes.FindField('PAS_DATEPLANSOC').AsDateTime));
     If (QAttes.FindField('PAS_DATEPREAVISD2').AsDateTime <> IDate1900) then
         SetControlText('DATCTAPP', DateToStr(QAttes.FindField('PAS_DATEPREAVISD2').AsDateTime));
     If (QAttes.FindField('PAS_DATEPREAVISF2').AsDateTime <> IDate1900) then
         SetControlText('DATEMBOCCAS', DateToStr(QAttes.FindField('PAS_DATEPREAVISF2').AsDateTime));
     If (QAttes.FindField('PAS_DATEPREAVISD3').AsDateTime <> IDate1900) then
         SetControlText('DATDEBOCCAS', DateToStr(QAttes.FindField('PAS_DATEPREAVISD3').AsDateTime));
     If (QAttes.FindField('PAS_DATEPREAVISF3').AsDateTime <> IDate1900) then
         SetControlText('DATFINOCCAS', DateToStr(QAttes.FindField('PAS_DATEPREAVISF3').AsDateTime));
     If (QAttes.FindField('PAS_INDPREAVIS').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC001'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDPREAVIS').asFloat;
     End;
     If (QAttes.FindField('PAS_INDCPMONTANT').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC002'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDCPMONTANT').asFloat;
     End;
     If (QAttes.FindField('PAS_INDINHER').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC003'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDINHER').asFloat;
     End;
     If (QAttes.FindField('PAS_INDLEGM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC004'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDLEGM').asFloat;
     End;
     If (QAttes.FindField('PAS_INDCDDM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC005'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDCDDM').asFloat;
     End;
     If (QAttes.FindField('PAS_INDCNEM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC006'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDCNEM').asFloat;
     End;
     If (QAttes.FindField('PAS_INDAUTREM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC007'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDAUTREM').asFloat;
     End;
     If (QAttes.FindField('PAS_INDFINMISSM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTC008'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDFINMISSM').asFloat;
     End;
     If (QAttes.FindField('PAS_HORHEBENT').asFloat <> 0) Then
     Begin
        Edit := THEdit(GetControl('HOCCAS'));
        If (Edit <> nil) Then Edit.Text := FloatToStr(QAttes.FindField('PAS_HORHEBENT').asFloat);
     End;
     If (QAttes.FindField('PAS_INDRETRAITEM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTD001'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDRETRAITEM').asFloat;
     End;
     If (QAttes.FindField('PAS_INDSINISTREM').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('MNTD002'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDSINISTREM').asFloat;
     End;
     SetControlText('CKINTERTRAV',QAttes.FindField('PAS_RECLASS').asstring);
     SetControlText('CKEMBACC',QAttes.FindField('PAS_PLANSOC').asstring);

     STrlu := QAttes.FindField('PAS_MAINTIEN').asstring;
     If STrlu = 'X' Then SetControlText('MAINTIENOUI','X');
     If STrlu = '-' Then SetControlText('MAINTIENNON','X');

     STrlu := QAttes.FindField('PAS_TYPMAINTIEN').asstring;
     If STrlu = 'INT' Then SetControlText('MAINTIENINT','X');
     If STrlu = 'PAR' Then SetControlText('MAINTIENPART','X');

     If (QAttes.FindField('PAS_MONTANT').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('SUBMT'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_MONTANT').asFloat;
     End;
     If (QAttes.FindField('PAS_SUBDEBUT').AsDateTime <> IDate1900) then
         SetControlText('DATEDEBSUBR', DateToStr(QAttes.FindField('PAS_SUBDEBUT').AsDateTime));
     If (QAttes.FindField('PAS_SUBFIN').AsDateTime <> IDate1900) then
         SetControlText('DATEFINSUBR', DateToStr(QAttes.FindField('PAS_SUBFIN').AsDateTime));

     SetControlText('DENOMINATION',QAttes.FindField('PAS_SUBCOMPTE').asstring);
     SetControlText('LIEUFAIT',QAttes.FindField('PAS_DECLARLIEU').asstring);
     If (QAttes.FindField('PAS_DECLARDATE').AsDateTime <> IDate1900) then
         SetControlText('DATEFAIT', DateToStr(QAttes.FindField('PAS_DECLARDATE').AsDateTime));
     SetControlText('NOMSIGNE',QAttes.FindField('PAS_DECLARPERS').asstring);
     SetControlText('QUALITESIGNE',QAttes.FindField('PAS_DECLARQUAL').asstring);
  End;
  Ferme(QAttes);
  //Init Tableau
  QAttes := OpenSql('SELECT PAL_SALARIE,PAL_TYPEATTEST,PAL_ORDRE,PAL_MOIS,' +
     'PAL_ECHEANCE,PAL_DATEDEBUT,PAL_DATEFIN,PAL_NBHEURES,PAL_SALAIRE,PAL_AVANTNAT,PAL_INDEMNITE,' +
     'PAL_SOUMIS,PAL_DEDSUP,PAL_PARTSAL4,PAL_VERSEMENT,PAL_DEBPERVER,PAL_FINPERVER,PAL_MONTANTVER,' +
     'PAL_PARTSAL12,PAL_MOTIF,PAL_DEBARRET,PAL_FINARRET ' +
     'FROM ATTSALAIRES WHERE PAL_SALARIE="' + Matricule + '" AND PAL_TYPEATTEST="MSA" ' +
     'And PAL_ORDRE = '+Ordre+' ORDER BY PAL_MOIS', True);

  i := 0;
  While not QAttes.Eof do
  Begin
     i := i + 1;
     If QAttes.FindField('PAL_ECHEANCE').AsDateTime <> IDate1900 Then
     Begin
        Edit := THEdit(GetControl('DATEECHEANCE' + IntToStr(i)));
        if (Edit <> nil) then Edit.Text := DateToStr(QAttes.FindField('PAL_ECHEANCE').AsDateTime);
        Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
        if (Edit <> nil) and (QAttes.FindField('PAL_DATEDEBUT').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_DATEDEBUT').AsDateTime);
        Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
        if (Edit <> nil) and (QAttes.FindField('PAL_DATEFIN').AsDateTime <> IDate1900) then
           Edit.text := DateToStr(QAttes.FindField('PAL_DATEFIN').AsDateTime);
        Edit := THEdit(GetControl('HTRAV' + IntToStr(i)));
        if (Edit <> nil) then Edit.text := FloatToStr(QAttes.FindField('PAL_NBHEURES').asFloat);
        NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_SALAIRE').asFloat;
        NumEdit := THNumEdit(GetControl('AVANT' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_AVANTNAT').asFloat;
        NumEdit := THNumEdit(GetControl('INDEM' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_INDEMNITE').asFloat;
        NumEdit := THNumEdit(GetControl('SOUMIS' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_SOUMIS').asFloat;
        NumEdit := THNumEdit(GetControl('DEDSUP' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_DEDSUP').asFloat;
        NumEdit := THNumEdit(GetControl('PARTSAL' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_PARTSAL4').asFloat;
     End;
     If QAttes.FindField('PAL_VERSEMENT').AsDateTime <> IDate1900 Then
     Begin
        Edit := THEdit(GetControl('DATEVER' + IntToStr(i)));
        if (Edit <> nil) then Edit.Text := DateToStr(QAttes.FindField('PAL_VERSEMENT').AsDateTime);
        Edit := THEdit(GetControl('DATEDEBVER' + IntToStr(i)));
        if (Edit <> nil) and (QAttes.FindField('PAL_DEBPERVER').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_DEBPERVER').AsDateTime);
        Edit := THEdit(GetControl('DATEFINVER' + IntToStr(i)));
        if (Edit <> nil) and (QAttes.FindField('PAL_FINPERVER').AsDateTime <> IDate1900) then
           Edit.text := DateToStr(QAttes.FindField('PAL_FINPERVER').AsDateTime);
        NumEdit := THNumEdit(GetControl('BRUTVER' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_MONTANTVER').asFloat;
        NumEdit := THNumEdit(GetControl('PARTSALB' + IntToStr(i)));
        if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_PARTSAL12').asFloat;
     End;
     If Trim(QAttes.FindField('PAL_MOTIF').asstring) <> '' Then
     Begin
        Edit := THEdit(GetControl('MOTIFARRET' + IntToStr(i)));
        if Edit <> nil then Edit.text := QAttes.FindField('PAL_MOTIF').asstring;
        Edit := THEdit(GetControl('DATEDEBARRET' + IntToStr(i)));
        if (Edit <> nil) and (QAttes.FindField('PAL_DEBARRET').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_DEBARRET').AsDateTime);
        Edit := THEdit(GetControl('DATEFINARRET' + IntToStr(i)));
        if (Edit <> nil) and (QAttes.FindField('PAL_FINARRET').AsDateTime <> IDate1900) then
           Edit.text := DateToStr(QAttes.FindField('PAL_FINARRET').AsDateTime);
     End;
     QAttes.next;
  End;
  Ferme(QAttes);
  Loaded := False;
End;

Procedure TOF_PGACCTRAVAIL_MSA.InitPages(Var Complet : Boolean);
Var
  i : Integer;
  NumEdit: THNumEdit;
Begin
  If Complet = TRUE Then
  Begin
     DerjourTravSaisie :='';
     SetControlText('CKVICTOUI','-');
     SetControlText('CKVICTNON','-');
     SetControlText('DATEACCTRAV','01/01/1900');
     SetControlText('DATEDERTRAV','01/01/1900');
     SetControlText('HEUREACCTRAV','00:00');
     SetControlText('DATEREPTRAV','01/01/1900');
     SetControlText('CKMOTIFACCTRAV','-');
     SetControlText('CKMOTIFMALPROF','-');
     SetControlText('CKNONREPRIS','-');

     SetControlText('DATEACCTRAV1','');
     SetControlText('DATEDERTRAV1','');
     SetControlText('HEUREACCTRAV1','');
     SetControlText('DATEREPTRAV1','');
     SetControlText('DATEMBOCCAS1','');
     SetControlText('DATDEBACTIV1','');
     SetControlText('DATCTAPP1','');

     For i := 1 to 10 do
     Begin
        SetControlText('DATEVER' + IntToStr(i),'01/01/1900');
        SetControlText('DATEDEBVER' + IntToStr(i),'01/01/1900');
        SetControlText('DATEFINVER' + IntToStr(i),'01/01/1900');
        SetControlText('BRUTVER' + IntToStr(i),'');
        SetControlText('PARTSALB' + IntToStr(i),'');
        NumEdit := THNumEdit(GetControl('BRUTVER' + IntToStr(i)));
        If NumEdit <> nil Then NumEdit.value := 0;
        NumEdit := THNumEdit(GetControl('PARTSALB' + IntToStr(i)));
        If NumEdit <> nil Then NumEdit.value := 0;
     End;

     SetControlText('DATDEBPART','01/01/1900');
     SetControlText('DATFINPART','01/01/1900');
     SetControlText('DATDEBACTIV','01/01/1900');
     SetControlText('DATEMBOCCAS','01/01/1900');
     SetControlText('DATDEBOCCAS','01/01/1900');
     SetControlText('DATFINOCCAS','01/01/1900');
     SetControlText('DATCTAPP','01/01/1900');
     SetControlText('HOCCAS','');
     SetControlText('MNTC001','');
     SetControlText('MNTC002','');
     SetControlText('MNTC003','');
     SetControlText('MNTC004','');
     SetControlText('MNTC005','');
     SetControlText('MNTC006','');
     SetControlText('MNTC007','');
     SetControlText('MNTC008','');
     SetControlText('CKINTERTRAV','-');
     SetControlText('CKEMBACC','-');
     SetControlText('MNTD001','');
     SetControlText('MNTD002','');
     NumEdit := THNumEdit(GetControl('MNTC001'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC002'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC003'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC004'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC005'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC006'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC007'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTC008'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTD001'));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('MNTD002'));
     If NumEdit <> nil Then NumEdit.value := 0;

     For i := 1 to 3 do
     Begin
        SetControlText('MOTIFARRET' + IntToStr(i),'');
        SetControlText('DATEDEBARRET' + IntToStr(i),'01/01/1900');
        SetControlText('DATEFINARRET' + IntToStr(i),'01/01/1900');
     End;

     SetControlText('MAINTIENOUI','-');
     SetControlText('MAINTIENNON','-');
     SetControlText('MAINTIENINT','-');
     SetControlText('MAINTIENPART','-');
     SetControlText('SUBMT','');
     SetControlText('DATEDEBSUBR','01/01/1900');
     SetControlText('DATEFINSUBR','01/01/1900');
     SetControlText('DENOMINATION', '');
     SetControlEnabled('BANQUE', False);
     SetControlText('BANQUE', '');
     SetControlEnabled('GUICHET', False);
     SetControlText('GUICHET', '');
     SetControlEnabled('COMPTE', False);
     SetControlText('COMPTE', '');
     SetControlEnabled('CLE', False);
     SetControlText('CLE', '');
     SetControlText('LIEUFAIT','');
     SetControlText('DATEFAIT','01/01/1900');
     SetControlText('NOMSIGNE','');
     SetControlText('QUALITESIGNE','');
     NumEdit := THNumEdit(GetControl('SUBMT'));
     If NumEdit <> nil Then NumEdit.value := 0;
  End;
  For i := 1 to 4 do
  Begin
     SetControlText('DATEECHEANCE' + IntToStr(i),'01/01/1900');
     SetControlText('DATEDEBPER' + IntToStr(i),'01/01/1900');
     SetControlText('DATEFINPER' + IntToStr(i),'01/01/1900');
     SetControlText('HTRAV' + IntToStr(i),'');
     SetControlText('MTSAL' + IntToStr(i),'');
     SetControlText('AVANT' + IntToStr(i),'');
     SetControlText('INDEM' + IntToStr(i),'');
     SetControlText('SOUMIS' + IntToStr(i),'');
     SetControlText('DEDSUP' + IntToStr(i),'');
     SetControlText('PARTSAL' + IntToStr(i),'');

     NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('AVANT' + IntToStr(i)));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('INDEM' + IntToStr(i)));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('SOUMIS' + IntToStr(i)));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('DEDSUP' + IntToStr(i)));
     If NumEdit <> nil Then NumEdit.value := 0;
     NumEdit := THNumEdit(GetControl('PARTSAL' + IntToStr(i)));
     If NumEdit <> nil Then NumEdit.value := 0;
  End;
End;

Procedure TOF_PGACCTRAVAIL_MSA.Imprimer(Sender: TObject);
var
  Pages : TPageControl;
  StPages, Requete : String;
Begin
  Pages := TPageControl(GetControl('PAGES'));
  Requete := '';
  StPages := AglGetCriteres (Pages, FALSE);
  LanceEtat('E','PAT','AT1',True,False,False,Pages,Requete,'',False,0,StPages);
  LanceEtat('E','PAT','AT2',True,False,False,Pages,Requete,'',False,0,StPages);
End;

Procedure TOF_PGACCTRAVAIL_MSA.OnLoad;
Var
  Page     : TPageControl;
Begin
  Loaded := False;
  Page := TPageControl(GetControl('PAGES'));
  If Page <> nil Then Page.ActivePageIndex := 0;
End;

Procedure TOF_PGACCTRAVAIL_MSA.OnNew;
Var
  Page   : TPageControl;
  AFaire : Boolean;
Begin
  AFaire := True;
  InitPages(AFaire);
  Page := TPageControl(GetControl('PAGES'));
  If Page <> nil Then Page.ActivePageIndex := 0;
  SetFocusControl('DATEACCTRAV');
End;

Procedure TOF_PGACCTRAVAIL_MSA.DateElipsisclick(Sender: TObject);
var
  key: char;
Begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
End;

Function TOF_PGACCTRAVAIL_MSA.FImpr(St : string): string;
Var
  StFormat: string;
  i,j,Lg,Esp : integer;
Begin
  If St = '' Then Begin result:=''; Exit; End;
  If St = '01/01/1900' Then Begin result:=''; Exit; End;
  If St = '00:00' Then Begin result:=''; Exit; End;
  Lg := Length(St);
  StFormat := Stringofchar (' ',Lg+5*(Lg-1));
  j:=1;
  Esp := 5;
  For i:=1 to Lg do
    If (St[i]<>'/') And (St[i]<>':') then
       Begin
         StFormat[j] := St[i];
         j := j+Esp;
       End;
  result:=Trim(StFormat);
End;

initialization
  registerclasses([TOF_PGACCTRAVAIL_MSA]);
end.

