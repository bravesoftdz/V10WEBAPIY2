{***********UNITE*************************************************
Auteur  ...... : PAIE PGI / RMA
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Attestation maladie MSA
Mots clefs ... : PAIE;MALADIE
*****************************************************************
PT1 30/05/2007 FC V_72 FQ 12925 Gestion du type de lien de parenté dans les personnes à charge
PT2 05/06/2007 RM V_72 FQ 14323 Gestion de la Loop sur Fiche individuelle et Fiche absence
                       FQ 14339 / FQ 14322 / FQ 14325 Ergonomie + amélioration
}

unit UTOFPGEditMaladieMSA;

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
  PgOutilsTreso, Commun, LookUp;

type
  TOF_PGMALADIE_MSA = class(TOF)
  Private
    Loaded, Ctrl, subrogation : Boolean;
    Matricule,GblEtab : string;
    Mode: string;
    Salaires : THGRID;
    ColEnCours, LigEnCours : integer;

    Procedure DateElipsisclick(Sender: TObject);
    Procedure InitSalarie;
    Procedure InitAttes;
    Procedure AffectSalaire;
    Procedure AffectDateArret(Sender: TObject);
    Procedure AffichDeclarant (Sender: TObject);
    Procedure Ck200Click(Sender: TObject);
    Procedure Ck800Click(Sender: TObject);
    Procedure EditExit(Sender: TObject);
    Procedure NumEditExit(Sender: TObject);
    Procedure CheckExit(Sender: TObject);
    Procedure CellExit(Sender: TObject; var ACol,ARow: Integer;
                        var Cancel: Boolean);
    Procedure GridExit(Sender: TObject);
    Procedure SalairesEnter(Sender: TObject);
    Procedure SALAIREDblClick(Sender: TObject);
    Procedure Controle_periode(var ValPer1,Valper2 : String);
    Procedure Imprimer(Sender: TObject);
    Procedure UpdateTable(Sender: TObject);
    Procedure InitPages(Var Complet : Boolean);
    Function  SaisieNonValide : Boolean;
    Function  FImpr(St ,StT : string): string;  //PT2
  Public
    Procedure OnArgument(Arguments: string); override;
    Procedure OnLoad; override;
    Procedure OnNew; override;
  End;

implementation

uses PgEditOutils, EntPaie;

procedure TOF_PGMALADIE_MSA.OnArgument(Arguments: string);
var
  Edit     : THEdit;
  NumEdit  : THNumEdit;
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
  Check := TCheckBox(GetControl('CKMOTIFMAL'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKMOTIFMAT'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKMOTIFPAT'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKMOTIFJOURMAT'));
  if Check <> nil then Check.OnClick := CheckExit;

  Edit := THEdit(GetControl('DATEARRET'));
  if Edit <> nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATEDERTRAV'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := AffectDateArret;
  end;
  Edit := THEdit(GetControl('DATEDEBPATER'));
  If Edit <> Nil then //PT2
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATEREPTRAV'));
  If Edit <> Nil then //PT2
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATESUSPCRT'));
  If Edit <> Nil then //PT2
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATEREPRPATIEL'));
  If Edit <> Nil then //PT2
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATEDEBPER'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;
  Edit := THEdit(GetControl('DATEFINPER'));
  If Edit <> Nil then
  begin
    Edit.OnElipsisClick := DateElipsisclick;
    Edit.OnExit := EditExit;
  end;

  Check := TCheckBox(GetControl('CKMOTIFPER'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CK200'));
  if Check <> nil then Check.OnClick := Ck200Click;
  Check := TCheckBox(GetControl('CK800'));
  if Check <> nil then Check.OnClick := Ck800Click;
  Check := TCheckBox(GetControl('CKCASGEN'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKCASPART'));
  if Check <> nil then Check.OnClick := CheckExit;
  NumEdit := THNumEdit(GetControl('MTCOT200'));
  If NumEdit <> Nil then NumEdit.OnExit := NumEditExit;
  NumEdit := THNumEdit(GetControl('MTCOT800'));
  If NumEdit <> Nil then NumEdit.OnExit := NumEditExit;

  //Evenement Volet 2
  Salaires:= THGRID (GetControl ('SALAIRE_REF'));
  If Salaires <> NIL then
  begin
  // valeurs des colonnes numériques cadrées à droite de la cellule
    Salaires.ColAligns[0]:= taCenter;
    Salaires.ColAligns[1]:= taCenter;
    Salaires.ColAligns[2]:= taRightJustify;
    Salaires.ColAligns[3]:= taRightJustify;
    Salaires.ColAligns[4]:= taRightJustify;
    Salaires.ColAligns[5]:= taCenter;
    Salaires.ColAligns[6]:= taRightJustify;
    Salaires.ColAligns[7]:= taRightJustify;
    Salaires.ColAligns[8]:= taRightJustify;

    Salaires.HideSelectedWhenInactive:= true;
    Salaires.OnCellExit:=CellExit;
    Salaires.OnExit:= GridExit;
    Salaires.OnDblClick:= SALAIREDblClick;
    Salaires.OnElipsisClick:= SALAIREDblClick;
    Salaires.OnClick:= SalairesEnter;

    Salaires.ColTypes[0]:= 'D';
    Salaires.colformats[0]:= ShortDateFormat;
    Salaires.ColTypes[1]:= 'D';
    Salaires.colformats[1]:= ShortDateFormat;
  End;
  For i := 1 to 3 do
  Begin
    Edit := THEdit(GetControl('DEBPERTHE' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('FINPERTHE' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('DEBPERADH' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('FINPERADH' + IntToStr(i)));
    If Edit <> nil Then
    Begin
      Edit.OnElipsisClick := DateElipsisclick;
      Edit.OnExit := EditExit;
    End;
    Edit := THEdit(GetControl('DATEADH' + IntToStr(i)));
    If Edit <> nil Then Edit.OnElipsisClick := DateElipsisclick;
  End;

  Edit := THEdit(GetControl('DATEFAIT'));
  If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
  Check := TCheckBox(GetControl('CKSUBINT'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKSUBPART'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKSUBOUI'));
  if Check <> nil then Check.OnClick := CheckExit;
  Check := TCheckBox(GetControl('CKSUBNON'));
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

  St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GblEtab+'%") '+
      ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MAL%" )  ' ;
  SetControlProperty('DECLARANT', 'Plus' ,St );
  Edit := THEdit(GetControl('DECLARANT'));
  if Edit <> nil then Edit.OnExit := AffichDeclarant;

  If Mode = 'MODIFICATION' Then
  Begin
      If (ExisteSQL('SELECT PAS_SALARIE FROM ATTESTATIONS WHERE PAS_SALARIE="' + Matricule + '"' +
         'And PAS_TYPEATTEST = "MSA" And PAS_ASSEDICCAISSE = "MAL"') = TRUE) Then
          // Je me sert de PAS_ASSEDICCAISSE pour MAL=Attestation Maladie et ACT = Attestation AT

          InitAttes
      Else
          Mode := 'CREATION';
  End;
  If Mode = 'CREATION' Then
  Begin
      QAttes:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                    'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GblEtab+'%") '+
                    'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MAL%" )  '+
                    'ORDER BY PDA_ETABLISSEMENT DESC',True);
      If Not QAttes.eof then
      Begin
           SetControlText('DECLARANT' ,QAttes.FindField('PDA_DECLARANTATTES').AsString);
           AffichDeclarant(nil);
           SetControlText('DATEFAIT', DateToStr(Date));
      End;
      Ferme(QAttes);
  End;
End;

Procedure TOF_PGMALADIE_MSA.InitSalarie;
var
  QSal: TQuery;
  Defaut: THLabel;
  NbEnfant : String;
Begin
  subrogation := False;
  QSal := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_NOMJF,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,' +
       'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,' +
       'PSA_DATENAISSANCE,PSA_SEXE,' +
       'ET_ETABLISSEMENT,ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,' +
       'ET_TELEPHONE,ET_SIRET,ETB_ETABLISSEMENT,ETB_NUMMSA,ETB_SUBROGATION '+
       'FROM SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
       'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
       'WHERE PSA_SALARIE="' + Matricule + '"', TRUE);

  If not QSal.EOF Then
  Begin
    GblEtab := QSal.FindField('ET_ETABLISSEMENT').asstring;
    Defaut  := THLabel(GetControl('ETLIBELLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_LIBELLE').asstring;
    Defaut := THLabel(GetControl('ETADRESSE1'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE1').asstring;
    Defaut := THLabel(GetControl('ETADRESSE23'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE2').asstring + ' ' + QSal.FindField('ET_ADRESSE3').asstring;

    Defaut := THLabel(GetControl('ETCP'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring, 'STR', 4);
    Defaut := THLabel(GetControl('ETVILLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('ET_VILLE').asstring;
    Defaut := THLabel(GetControl('ETTELEPHONE'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_TELEPHONE').asstring, 'STR', 2);
    Defaut := THLabel(GetControl('ETSIRET'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_SIRET').asstring, 'STR', 4);

    if (QSal.FindField('PSA_SEXE').asstring = 'M') Or (Trim(QSal.FindField('PSA_NOMJF').asstring) = '') then
       SetControlText('PSALIBELLE', QSal.FindField('PSA_LIBELLE').asstring)
    else
       SetControlText('PSALIBELLE', QSal.FindField('PSA_NOMJF').asstring + '    (épouse ' + QSal.FindField('PSA_LIBELLE').asstring + ')');

    SetControlText('PSAPRENOM', QSal.FindField('PSA_PRENOM').asstring);

    Defaut := THLabel(GetControl('PSADATENAISS'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_DATENAISSANCE').asstring;
    Defaut := THLabel(GetControl('PSAADRESSE1'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_ADRESSE1').asstring;
    Defaut := THLabel(GetControl('PSAADRESSE23'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_ADRESSE2').asstring + '  ' + QSal.FindField('PSA_ADRESSE3').asstring;
    Defaut := THLabel(GetControl('PSAEMPLOI'));
    if defaut <> nil then Defaut.caption := RechDom('PGLIBEMPLOI', QSal.FindField('PSA_LIBELLEEMPLOI').asstring, False);
    Defaut := THLabel(GetControl('PSACP'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring, 'STR', 4);
    Defaut := THLabel(GetControl('PSAVILLE'));
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_VILLE').asstring;
    Defaut := THLabel(GetControl('PSASECU1'));
    if (defaut <> nil) then defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 1, 13), 'STR', 4);
    Defaut := THLabel(GetControl('PSASECU2'));
    if (defaut <> nil) then Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 14, 2), 'STR', 3);
    Defaut := THLabel(GetControl('PSASALARIE'));                                    //PT2
    if defaut <> nil then Defaut.caption := QSal.FindField('PSA_SALARIE').asstring; //PT2
    subrogation := QSal.FindField('ETB_SUBROGATION').Asstring = 'X';
  End;
  Ferme(QSal);
  NbEnfant := '0';
  QSal := OpenSql('SELECT COUNT(PEF_ENFANT) FROM ENFANTSALARIE ' +
       'WHERE PEF_SALARIE="' + Matricule + '" AND PEF_ACHARGE = "X" AND PEF_TYPEPARENTAL="001"', TRUE); //PT1

  If not QSal.EOF Then
  Begin
    NbEnfant := QSal.Fields[0].asstring;
  End;
  Ferme(QSal);
  SetControlText('NBENFANT', NbEnfant);
End;

Procedure TOF_PGMALADIE_MSA.AffectSalaire;
Var
  QRech : TQuery;
  Edit : THedit;
  DateD, DerJour,DateDeb, DateFin : TDateTime;
  i, j : integer;
  MontZone : THNumEdit;
  CasGen, CasPart, MotifMal, MotifMat, MotifPat, MotifJMat, Ck : TCheckBox;
  Tob_Montant, TMont : Tob;
  Trouve : boolean;
begin
  if Loaded = True then exit;
  MotifMal := TCheckBox(GetControl('CKMOTIFMAL'));
  MotifMat := TCheckBox(GetControl('CKMOTIFMAT'));
  MotifPat := TCheckBox(GetControl('CKMOTIFPAT'));
  MotifJMat:= TCheckBox(GetControl('CKMOTIFJOURMAT'));
  CasGen   := TCheckBox(GetControl('CKCASGEN'));
  CasPart  := TCheckBox(GetControl('CKCASPART'));
  Edit     := THEdit(GetControl('DATEARRET'));

  If (Edit <> nil) And (CasGen <> nil) And (CasPart <> nil) And (MotifMal <> nil) And (MotifMat <> nil)
     And (MotifPat <> nil) And (MotifJMat <> nil) Then
  Begin
     If (MotifMal.checked = False) And (MotifMat.checked = False) And (MotifPat.checked = False) And (MotifJMat.checked = False) Then
        Exit;
     If (CasGen.checked = False) and (CasPart.checked = False) Then Exit;
     If Not IsValidDate(Edit.text) Then Exit;

     Trouve := False;
     InitPages(Trouve);
     DerJour := StrToDate(Edit.text);
     DateFin := DebutdeMois(DerJour)-1;

     If (CasGen.Checked = True) Then
     Begin
        MontZone := THNumEdit(GetControl('MTCOT200'));
        DateDeb := DebutDeMois(PlusMois(DebutDeMois(DerJour)-1, -5));
     End
     Else
     Begin
        MontZone := THNumEdit(GetControl('MTCOT800'));
        DateDeb := DebutDeMois(PlusMois(DebutDeMois(DerJour)-1, -11));
     End;
     If MontZone <> nil then
     Begin
        QRech := OpenSql('SELECT SUM(PHB_MTSALARIAL) AS CRDS FROM HISTOBULLETIN ' +
          'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
          'WHERE PCT_SOUMISMALAD="X"  AND PHB_SALARIE="' + Matricule + '"  ' +
          'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ', True);

        If not QRech.EOF then
        Begin
           MontZone.Value := QRech.FindField('CRDS').asFloat;
           SetControlText('DATEDEBPER',DateToStr(DateDeb));
           SetControlText('DATEFINPER',DateToStr(DateFin));
        End;
        Ferme(Qrech);
     End;
     //Affectation "PLus de 200h"
     If (CasGen.Checked = True) then
     Begin
       Ck := TCheckBox(GetControl('CK200'));
       DateDeb := DebutDeMois(PlusMois(DerJour, -2));
     End
     Else
     Begin
       Ck := TCheckBox(GetControl('CK800'));
       DateDeb := DebutDeMois(PlusMois(DerJour, -11));
     End;
     If Ck <> nil then
     Begin
       QRech := OpenSql('SELECT SUM(PPU_CHEURESTRAV) AS HEURE FROM PAIEENCOURS ' +
          'WHERE PPU_SALARIE="' + Matricule + '" ' +
          'AND PPU_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '"', True);

       If not QRech.EOF then //PORTAGECWAS
       Begin
          If (CasGen.Checked = True) then
             Ck.Checked := (QRech.FindField('HEURE').asFloat >= 200)
          Else
             Ck.Checked := (QRech.FindField('HEURE').asFloat >= 800);
       End;
       Ferme(Qrech);
     End;
     //Remplissage du tableau
     Tob_Montant := Tob.create('Histo_Bulletin', nil, -1);
     If (CasGen.Checked = True) then
        DateDeb := DebutDeMois(PlusMois(DerJour, -3))
     Else
        DateDeb := DebutDeMois(PlusMois(DerJour, -12));

     QRech := OpenSql('SELECT DISTINCT PPU_DATEDEBUT DATEDEBUT,PPU_DATEFIN DATEFIN,' +
        'SUM(PPU_CBASESS) AS CBASESS,' +
        'SUM(PPU_OCBASESS) AS OCBASESS,' +
        'SUM(PPU_CCOUTSALARIE) AS CCOUTSALARIE,' +
        'SUM(PPU_OCCOUTSALARIE) AS OCCOUTSALARIE ' +
        'FROM PAIEENCOURS ' +
        'WHERE  PPU_SALARIE="' + Matricule + '" ' +
        'AND PPU_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
        'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN ORDER BY PPU_DATEDEBUT,PPU_DATEFIN', TRUE);

     Tob_Montant.LoadDetailDB('Histo_Bulletin', '', '', QRech, False);
     Ferme(QRech);
     TMont := Tob_montant.FindFirst([''], [''], False);
     While TMont <> nil do
     Begin
        TMont.AddChampSup('CRDS', False);
        TMont.AddChampSup('OCRDS', False);
        TMont.PutValue('CRDS', 0);
        TMont.PutValue('OCRDS', 0);
        TMont := Tob_montant.FindNext([''], [''], False);
     End;
     QRech := OpenSql('SELECT DISTINCT PHB_DATEDEBUT DATEDEBUT,PHB_DATEFIN DATEFIN,' +
        'SUM(PHB_MTSALARIAL) AS CRDS,' +
        'SUM(PHB_OMTSALARIAL) AS OCRDS FROM HISTOBULLETIN ' +
        'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
        'WHERE PCT_ETUDEDROIT="X"  AND PHB_SALARIE="' + Matricule + '" ' +
        'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' + //PT2-2   Ajout du <=
        'GROUP BY PHB_DATEDEBUT,PHB_DATEFIN ORDER BY PHB_DATEDEBUT,PHB_DATEFIN', TRUE);
     Trouve := False;
     While not QRech.eof do
     Begin
        TMont := Tob_montant.FindFirst([''], [''], False);
        While TMont <> nil do
        Begin
          If (TMont.GetValue('DATEDEBUT') = QRech.FindField('DATEDEBUT').AsDateTime)
            and (TMont.GetValue('DATEFIN') = QRech.FindField('DATEFIN').AsDateTime) then
          Begin
            Trouve := True;
            TMont.PutValue('CRDS', TMont.GetValue('CRDS') + QRech.FindField('CRDS').AsFloat);
            TMont.PutValue('OCRDS', TMont.GetValue('OCRDS') + QRech.FindField('OCRDS').AsFloat);
          End;
          TMont := Tob_montant.FindNext([''], [''], False);
        End;
        If not Trouve then
        Begin
          TMont := Tob.create('La fille', tob_Montant, -1);
          If TMont <> nil then
          Begin
            TMont.PutValue('DATEDEBUT', QRech.FindField('DATEDEBUT').AsDateTime);
            TMont.PutValue('DATEFIN', QRech.FindField('DATEFIN').AsDateTime);
            TMont.PutValue('CRDS', QRech.FindField('CRDS').AsFloat);
            TMont.PutValue('OCRDS', QRech.FindField('OCRDS').AsFloat);
          End;
        End;
        Trouve := False;
        QRech.next;
     End;
     Ferme(QRech);
     j := 0;
     TMont := Tob_montant.FindFirst([''], [''], False);
     While TMont <> nil do
     Begin //B While
        i := 0;
        j := j + 1;
        Salaires.CellValues[i,j] := DateToStr(TMont.GetValue('DATEDEBUT'));
        i := i + 1;
        Salaires.CellValues[i,j] := DateToStr(TMont.GetValue('DATEFIN'));
        i := i + 1;
        DateD := TMont.GetValue('DATEDEBUT');

        If (MotifMal.Checked = True) Or (MotifJMat.checked = True) then
        Begin
           If (VH_Paie.PGDateBasculEuro <= DateD) Then
              Salaires.CellValues[i,j] := TMont.GetValue('CBASESS')
            Else
              Salaires.CellValues[i,j] := TMont.GetValue('OCBASESS');
        End;
        i := i + 1;
        If (MotifMat.Checked = True) Or (MotifPat.checked = True) then
        Begin
           If (VH_Paie.PGDateBasculEuro <= DateD) Then
              Salaires.CellValues[i,j] := TMont.GetValue('CBASESS')
            Else
              Salaires.CellValues[i,j] := TMont.GetValue('OCBASESS');

           i := i + 1;
           If (VH_Paie.PGDateBasculEuro <= DateD) Then
              Salaires.CellValues[i,j] := TMont.GetValue('CBASESS') - TMont.GetValue('CCOUTSALARIE') + TMont.GetValue('CRDS')
            Else
              Salaires.CellValues[i,j] := TMont.GetValue('OCBASESS') - TMont.GetValue('OCCOUTSALARIE') + TMont.GetValue('OCRDS');
        End;
        Tmont := Tob_montant.FindNext([''], [''], False); //QRech.Next;
     End; //End While
     If Tob_Montant <> nil then Tob_Montant.free;
  End;
End;

Procedure TOF_PGMALADIE_MSA.AffectDateArret(Sender: TObject);
Var               //PT2
  TXT : String;   //PT2
Begin
  If IsValidDate(GetControlText('DATEDERTRAV')) Then
  Begin
     SetControlText('DATEARRET', DateToStr(StrToDate(GetControlText('DATEDERTRAV')) + 1));
     TXT := FImpr(GetControlText('DATEDERTRAV'),'N');  //PT2
     SetControlText('DATEDERTRAV1', TXT);              //PT2
     AffectSalaire;
  End;
End;

//PT2 Ajout de la fonction FImpr ==>
Function  TOF_PGMALADIE_MSA.FImpr(St ,StT : string): string;
Var
  StFormat: string;
  i,j,Lg,Esp : integer;
Begin
  If St = '' Then Begin result:=''; Exit; End;
  If St = '01/01/1900' Then Begin result:=''; Exit; End;
  Lg := Length(St);
  StFormat := Stringofchar (' ',Lg+5*(Lg-1));
  j:=1;
  Esp := 3;
  For i:=1 to Lg do
     If (St[i]<>'/') then
     Begin
       StFormat[j] := St[i];
       j := j+Esp;
       If (i=2) then j := j+1;
       If (i=3) And (StT = 'C') then j := j+2;
       If (i=4) And (StT = 'C') then j := j+1;
       If (i=5) And (StT = 'N') then j := j+2;
       If (i=5) And (StT = 'C') then j := j+1;
       If (i=7) then j := j+1;
       If (i=8) And (StT = 'N') then j := j+2;
       If (i=9) And (StT = 'N') then j := j+1;
     End;
  result:=Trim(StFormat);
End; //PT2 <== FIN

Procedure TOF_PGMALADIE_MSA.AffichDeclarant(Sender: TObject);
Var
  St : String;
Begin
  If GetControlText('DECLARANT')='' Then exit;
  St := RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False);
  SetControlText('NOMSIGNE' ,RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
  SetControlText('LIEUFAIT'        ,RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
  St := RechDom('PGDECLARANTQUAL'  ,GetControlText('DECLARANT'),False);
  If St = 'AUT' Then SetControlText('QUALITESIGNE' ,RechDom('PGDECLARANTAUTRE' ,GetControlText('DECLARANT'),False))
                Else SetControlText('QUALITESIGNE' ,RechDom('PGQUALDECLARANT2' ,St,False));
End;

Procedure TOF_PGMALADIE_MSA.Ck200Click(Sender: TObject);
Var
  DateDeb, DateFin : TDateTime;
Begin
  IF GetControlText('CK200') = 'X' then
  Begin
     SetControlText('MTCOT200','');
     SetControlEnabled('MTCOT200',False);

     If IsValidDate(GetControlText('DATEARRET')) then
     Begin
       DateDeb := StrToDate(GetControlText('DATEARRET'));
       DateFin := DebutdeMois(DateDeb)-1;
       DateDeb := DebutDeMois(PlusMois(DebutDeMois(DateDeb)-1, -2));
       SetControlText('DATEDEBPER',DateToStr(DateDeb));
       SetControlText('DATEFINPER',DateToStr(DateFin));
     End;
  End
  Else
  Begin
     SetControlEnabled('MTCOT200',True);
     AffectSalaire;
  End;
End;

Procedure TOF_PGMALADIE_MSA.Ck800Click(Sender: TObject);
Var
  DateDeb, DateFin : TDateTime;
Begin
  IF GetControlText('CK800') = 'X' then
  Begin
     SetControlText('MTCOT800','');
     SetControlEnabled('MTCOT800',False);

     If IsValidDate(GetControlText('DATEARRET')) then
     Begin
       DateDeb := StrToDate(GetControlText('DATEARRET'));
       DateFin := DebutdeMois(DateDeb)-1;
       DateDeb := DebutDeMois(PlusMois(DebutDeMois(DateDeb)-1, -11));
       SetControlText('DATEDEBPER',DateToStr(DateDeb));
       SetControlText('DATEFINPER',DateToStr(DateFin));
     End;
  End
  Else
  Begin
     SetControlEnabled('MTCOT800',True);
     AffectSalaire;
  End;
End;

Procedure TOF_PGMALADIE_MSA.NumEditExit(Sender: TObject);
Begin
  If THNumEdit(Sender).Name = 'MTCOT200' Then
     IF (GetControlText('MTCOT200') <> '0,00') Then
         SetControlText('MTCOT800','');
  If THNumEdit(Sender).Name = 'MTCOT800' Then
     IF (GetControlText('MTCOT800') <> '0,00') Then
        SetControlText('MTCOT200','');
End;

Procedure TOF_PGMALADIE_MSA.EditExit(Sender: TObject);
var
  Valeur1, Valeur2 : String;
  TXT : String;   //PT2
Begin
  If THEdit(Sender).Name = 'DATEARRET' Then
     If IsValidDate(GetControlText('DATEARRET')) Then AffectSalaire;

  If THEdit(Sender).Name = 'DATEREPRPATIEL' Then
     If (IsValidDate(GetControlText('DATEREPRPATIEL'))) And (GetControlText('DATEREPRPATIEL')<> '01/01/1900') Then //PT2
     Begin  //PT2
        SetControlText('CKMOTIFPER','-');
        TXT := FImpr(GetControlText('DATEREPRPATIEL'),'N');  //PT2
        SetControlText('DATEREPRPATIEL1', TXT);              //PT2
     End;   //PT2

  If (THEdit(Sender).Name = 'DATEDEBPER') Or (THEdit(Sender).Name = 'DATEFINPER') Then
  Begin
     Valeur1 := GetControlText('DATEDEBPER');
     Valeur2 := GetControlText('DATEFINPER');
     //PT2 ===>
     If IsValidDate(Valeur1) Then
     Begin
        TXT := FImpr(Valeur1,'C');
        SetControlText('DATEDEBPER1', TXT);
     End;
     If IsValidDate(Valeur2) Then
     Begin
        TXT := FImpr(Valeur2,'C');
        SetControlText('DATEFINPER1', TXT);
     End;
     //PT2 <====
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DATEDEBPER');
  End;
  //PT2 ====>
  If THEdit(Sender).Name ='DATEDEBPATER' Then
  Begin
    TXT := FImpr(GetControlText('DATEDEBPATER'),'N');
    SetControlText('DATEDEBPATER1', TXT);
  End;
  If THEdit(Sender).Name ='DATEREPTRAV' Then
  Begin
    TXT := FImpr(GetControlText('DATEREPTRAV'),'N');
    SetControlText('DATEREPTRAV1', TXT);
  End;
  If THEdit(Sender).Name ='DATESUSPCRT' Then
  Begin
    TXT := FImpr(GetControlText('DATESUSPCRT'),'N');
    SetControlText('DATESUSPCRT1', TXT);
  End;
  //PT2 <=====
  If (THEdit(Sender).Name = 'DATEDEBSUBR') Or (THEdit(Sender).Name = 'DATEFINSUBR') Then
  Begin
     Valeur1 := GetControlText('DATEDEBSUBR');
     Valeur2 := GetControlText('DATEFINSUBR');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DATEDEBSUBR');
  End;

  If (THEdit(Sender).Name = 'DEBPERTHE1') Or (THEdit(Sender).Name = 'FINPERTHE1') Then
  Begin
     Valeur1 := GetControlText('DEBPERTHE1');
     Valeur2 := GetControlText('FINPERTHE1');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DEBPERTHE1');
  End;
  If (THEdit(Sender).Name = 'DEBPERTHE2') Or (THEdit(Sender).Name = 'FINPERTHE2') Then
  Begin
     Valeur1 := GetControlText('DEBPERTHE2');
     Valeur2 := GetControlText('FINPERTHE2');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DEBPERTHE2');
  End;
  If (THEdit(Sender).Name = 'DEBPERTHE3') Or (THEdit(Sender).Name = 'FINPERTHE3') Then
  Begin
     Valeur1 := GetControlText('DEBPERTHE3');
     Valeur2 := GetControlText('FINPERTHE3');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DEBPERTHE3');
  End;
  If (THEdit(Sender).Name = 'DEBPERADH1') Or (THEdit(Sender).Name = 'FINPERADH1') Then
  Begin
     Valeur1 := GetControlText('DEBPERADH1');
     Valeur2 := GetControlText('FINPERADH1');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DEBPERADH1');
  End;
  If (THEdit(Sender).Name = 'DEBPERADH2') Or (THEdit(Sender).Name = 'FINPERADH2') Then
  Begin
     Valeur1 := GetControlText('DEBPERADH2');
     Valeur2 := GetControlText('FINPERADH2');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DEBPERADH2');
  End;
  If (THEdit(Sender).Name = 'DEBPERADH3') Or (THEdit(Sender).Name = 'FINPERADH3') Then
  Begin
     Valeur1 := GetControlText('DEBPERADH3');
     Valeur2 := GetControlText('FINPERADH3');
     Controle_periode(Valeur1,Valeur2);
     If Ctrl = False Then SetFocusControl('DEBPERADH3');
  End;
End;

Procedure TOF_PGMALADIE_MSA.CheckExit(Sender: TObject);
Begin
  If TCheckBox(Sender).Name = 'CKMOTIFPER' Then
     IF GetControlText('CKMOTIFPER') = 'X' then SetControlText('DATEREPRPATIEL','');
  If TCheckBox(Sender).Name = 'CKSUBINT' Then
     IF GetControlText('CKSUBINT') = 'X' then SetControlText('CKSUBPART','-');
  If TCheckBox(Sender).Name = 'CKSUBPART' Then
     IF GetControlText('CKSUBPART') = 'X' then SetControlText('CKSUBINT','-');
  If TCheckBox(Sender).Name = 'CKSUBOUI' Then
     IF GetControlText('CKSUBOUI') = 'X' then SetControlText('CKSUBNON','-');
  If TCheckBox(Sender).Name = 'CKSUBNON' Then
     IF GetControlText('CKSUBNON') = 'X' then SetControlText('CKSUBOUI','-');
  If TCheckBox(Sender).Name = 'CKMOTIFMAL' Then
  Begin
     IF GetControlText('CKMOTIFMAL') = 'X' then
     Begin
        SetControlText('CKMOTIFMAT','-');
        SetControlText('CKMOTIFPAT','-');
        SetControlText('CKMOTIFJOURMAT','-');
        AffectSalaire;
     End;
  End;
  If TCheckBox(Sender).Name = 'CKMOTIFMAT' Then
  Begin
     IF GetControlText('CKMOTIFMAT') = 'X' then
     Begin
        SetControlText('CKMOTIFMAL','-');
        SetControlText('CKMOTIFPAT','-');
        SetControlText('CKMOTIFJOURMAT','-');
        AffectSalaire;
     End;
  End;
  If TCheckBox(Sender).Name = 'CKMOTIFPAT' Then
  Begin
     IF GetControlText('CKMOTIFPAT') = 'X' then
     Begin
        SetControlText('CKMOTIFMAL','-');
        SetControlText('CKMOTIFMAT','-');
        SetControlText('CKMOTIFJOURMAT','-');
        AffectSalaire;
     End;
  End;
  If TCheckBox(Sender).Name = 'CKMOTIFJOURMAT' Then
  Begin
     IF GetControlText('CKMOTIFJOURMAT') = 'X' then
     Begin
        SetControlText('CKMOTIFMAL','-');
        SetControlText('CKMOTIFMAT','-');
        SetControlText('CKMOTIFPAT','-');
        AffectSalaire;
     End;
  End;
  If TCheckBox(Sender).Name = 'CKCASGEN' Then
  Begin
     If GetControlText('CKCASGEN')= 'X' Then
     Begin
        SetControlText('CKCASPART','-');
        SetControlText('MTCOT800','');
        SetControlEnabled('MTCOT800',False);
        SetControlText('CK800','-');
        SetControlEnabled('CK800',False);
        SetControlEnabled('MTCOT200',True);
        SetControlEnabled('CK200',True);
        AffectSalaire;
     End;
  End;
  If TCheckBox(Sender).Name = 'CKCASPART' Then
  Begin
     If GetControlText('CKCASPART')= 'X' Then
     Begin
        SetControlText('CKCASGEN','-');
        SetControlText('MTCOT200','');
        SetControlEnabled('MTCOT200',False);
        SetControlText('CK200','-');
        SetControlEnabled('CK200',False);
        SetControlEnabled('MTCOT800',True);
        SetControlEnabled('CK800',True);
        AffectSalaire;
     End;
  End;
End;

Procedure TOF_PGMALADIE_MSA.CellExit (Sender: TObject; var ACol,ARow: Integer;
                                     var Cancel: Boolean);
var
  Valeur1, Valeur2 : String;
Begin
  If ((ACol=-1) And (ARow=-1)) Then Exit;
  Valeur1 := Salaires.CellValues[ACol,ARow];
  If Trim(Valeur1)='' Then Exit;

  If ((ACol = 0) or (ACol = 1)) And (ARow <> 0) Then
  Begin
     If Not IsValidDate(Valeur1) Then
     Begin
        PGIBox ('Date Invalide', TFVierge(Ecran).caption);
        Salaires.CellValues[ACol, ARow]:= '01/01/1900';
     End
     Else
     Begin
        If ACol = 0 Then
        Begin
           Valeur2 := Salaires.CellValues[1,ARow];
           Controle_periode(Valeur1,Valeur2);
           If Ctrl = False Then
           Begin
              Salaires.col := Acol;
              Salaires.row := Arow;
              Exit;
           End;
        End
        Else
        Begin
           Valeur2 := Salaires.CellValues[0,ARow];
           Controle_periode(Valeur2,Valeur1);
           If Ctrl = False Then
           Begin
              Salaires.col := Acol;
              Salaires.row := Arow;
              Exit;
           End;
        End;
     End;
  End;
  If ((ACol = 2) or (ACol = 3) or (ACol = 4) or (ACol = 8)) And (ARow<>0) then
     If Not IsNumeric(Valeur1) Then
         Salaires.CellValues[ACol, ARow]:= '0.00'
         Else Salaires.CellValues[ACol, ARow] := FormatFloat('#0.00',StrToFloat(Valeur1)); //PT2 '# ##0.00'

  If ((ACol = 6) or (ACol = 7)) And (ARow<>0) then
     If Not IsNumeric(Valeur1) Then
         Salaires.CellValues[ACol, ARow]:= '0';
End;

Procedure TOF_PGMALADIE_MSA.GridExit(Sender: TObject);
var
  Bool : boolean;
Begin
  CellExit (Sender, ColEnCours, LigEnCours, Bool);
  ColEnCours:= -1;
  LigEnCours:= -1;
End;

Procedure TOF_PGMALADIE_MSA.SalairesEnter(Sender: TObject);
Begin
  ColEnCours:= Salaires.Col;
  LigEnCours:= Salaires.Row;

  If ((ColEnCours = 0) or (ColEnCours = 1) Or (ColEnCours = 5)) And (LigEnCours <> 0) Then
    Salaires.ElipsisButton := TRUE
  Else
    Salaires.ElipsisButton := FALSE;
End;

Procedure TOF_PGMALADIE_MSA.SALAIREDblClick(Sender: TObject);
Var
  sWhere : string;

Begin
  If ((ColEnCours = 0) or (ColEnCours = 1)) And (LigEnCours <> 0) Then
     DateElipsisclick(Sender);
  If (ColEnCours = 5) And (LigEnCours <> 0) Then
  Begin
   sWhere:='##CO_PREDEFINI## CO_TYPE="PMT"';
   LookUpList (Salaires,'Motif des Absences Attestation','COMMUN','CO_CODE','CO_LIBELLE',sWhere,'CO_CODE',TRUE,-1);
  End;
End;

Procedure TOF_PGMALADIE_MSA.Controle_periode(var ValPer1,Valper2 : String);
Begin
  Ctrl := True;
  If (IsValidDate(ValPer1)) And (IsValidDate(ValPer2)) Then
  Begin
     If (StrToDate(ValPer1) > StrToDate(ValPer2)) And (ValPer2 <> '01/01/1900') Then  //PT2
     Begin
        PGIBox ('La période est invalide', TFVierge(Ecran).caption);
        Ctrl := False;
     End;
  End;
End;

Procedure TOF_PGMALADIE_MSA.UpdateTable(Sender: TObject);
var
  Check: TcheckBox;
  Edit: THEdit;
  NumEdit: THNumEdit;
  TComBo : ThvalComboBox;
  libelle, DerjourTrav, DateRepArret, MotifArret, Situation, NonRepris,DateDebPater : string;
  RepPartiel, Cas, Plus, DateDebPer, DateFinPer, LieuDecla, DateDecla, PersDecla, QualDecla : string;
  DateDebSubr, DateFinSubr, DateArret, NbEnfants, DateRepPart : string;
  StSal, MtSalRet, MtRegCotPer1, MtRegCotPer2, TravNuit, ExpoRisk, DateSupCtr, TMaint : string;
  Motif, NbHeures, NbHeureCompl, MoisSubr, CKMntsubro : string;
  MoisAdh, DateAdh, DebPerAdh, FinPerAdh, MtAdh : String;
  Montant, MntSubro : double;
  ordre, i, j : integer;
  QRech: TQuery;
  Tob_Mere, T: TOB;
begin
  If SaisieNonValide then Exit;
  ordre := 0;
  MotifArret := '';
  Check := TcheckBox(GetControl('CKMOTIFMAL'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MAL';
  Check := TcheckBox(GetControl('CKMOTIFMAT'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MAT';
  Check := TcheckBox(GetControl('CKMOTIFPAT'));
  if check <> nil then if Check.Checked = True then MotifArret := 'PAT';
  Check := TcheckBox(GetControl('CKMOTIFJOURMAT'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MJT';

  DerjourTrav := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if edit <> nil then if IsValidDate(Edit.text) then DerjourTrav := Edit.text;
  DateDebPater := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDEBPATER'));
  if edit <> nil then if IsValidDate(Edit.text) then DateDebPater := Edit.text;
  DateRepArret := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEREPTRAV'));
  if edit <> nil then if IsValidDate(Edit.text) then DateRepArret := Edit.text;
  Situation := '';
  TComBo := ThvalComboBox(GetControl('SITUATIONARRET'));
  if TComBo <> nil then if TComBo.value <> '' then Situation := TComBo.value;
  NbEnfants := '';
  Edit := THEdit(GetControl('NBENFANT'));
  if edit <> nil then if Edit.text <> '' then NbEnfants := Edit.text;
  DateRepPart := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEREPRPATIEL'));
  if edit <> nil then if IsValidDate(Edit.text) then DateRepPart := Edit.text;
  DateArret := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEARRET'));
  if edit <> nil then if IsValidDate(Edit.text) then DateArret := Edit.text;
  NonRepris := '-';
  Check := TcheckBox(GetControl('CKNONREPRIS'));
  if check <> nil then if Check.Checked = True then NonRepris := 'X';
  RepPartiel := '-';
  Check := TcheckBox(GetControl('CKMOTIFPER'));
  if check <> nil then if Check.Checked = True then RepPartiel := 'X';
  TravNuit := '-';
  Check := TcheckBox(GetControl('CKTRAVNUIT'));
  if check <> nil then if Check.Checked = True then TravNuit := 'X';
  ExpoRisk := '-';
  Check := TcheckBox(GetControl('CKEXPORISQ'));
  if check <> nil then if Check.Checked = True then ExpoRisk := 'X';
  DateSupCtr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATESUSPCRT'));
  if edit <> nil then if IsValidDate(Edit.text) then DateSupCtr := Edit.text;

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
  if edit <> nil then if IsValidDate(Edit.text) then DateDebPer := Edit.text;
  DateFinPer := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFINPER'));
  if edit <> nil then if IsValidDate(Edit.text) then DateFinPer := Edit.text;
  DateDebSubr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEDEBSUBR'));
  if edit <> nil then if IsValidDate(Edit.text) then DateDebSubr := Edit.text;
  DateFinSubr := DateToStr(Idate1900);
  Edit := THEdit(GetControl('DATEFINSUBR'));
  if edit <> nil then if IsValidDate(Edit.text) then DateFinSubr := Edit.text;
  TMaint := '';
  Check := TcheckBox(GetControl('CKSUBINT'));
  if check <> nil then if Check.Checked = True then TMaint := 'INT';
  Check := TcheckBox(GetControl('CKSUBPART'));
  if check <> nil then if Check.Checked = True then TMaint := 'PAR';
  MoisSubr := '';
  Edit := THEdit(GetControl('SUBMOIS'));
  if edit <> nil then if Edit.text <> '' then MoisSubr := Edit.text;
  MntSubro := 0;
  NumEdit := THNumEdit(GetControl('SUBMT'));
  if NumEdit <> nil then if NumEdit.Value > 0 then MntSubro := NumEdit.value;
  CKMntsubro := '';
  Check := TcheckBox(GetControl('CKSUBOUI'));
  if check <> nil then if Check.Checked = True then CKMntsubro := 'X';
  Check := TcheckBox(GetControl('CKSUBNON'));
  if check <> nil then if Check.Checked = True then CKMntsubro := '-';
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
      (TMaint = '') or (MoisSubr = '') or (MntSubro = 0) or (CKMntsubro = '')) Then

     PgiBox('Des informations concernant la subrogation sont absentes!','Subrogation');

  QRech := OpenSql('SELECT PAS_ORDRE FROM ATTESTATIONS ' +
    'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="MSA" AND PAS_ASSEDICCAISSE = "MAL"', True);

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
  T.PutValue('PAS_ASSEDICCAISSE', 'MAL');                 //Champ détourné
  T.PutValue('PAS_MOTIFARRET', MotifArret);
  T.PutValue('PAS_DERNIERJOUR', StrToDate(DerjourTrav));
  T.PutValue('PAS_DEBCHOMAGE', StrToDate(DateDebPater));  //Champ détourné
  T.PutValue('PAS_REPRISEARRET', StrToDate(DateRepArret));
  T.PutValue('PAS_SITUATION', Situation);
  T.PutValue('PAS_CONVNUMERO', NbEnfants);                //Champ détourné
  T.PutValue('PAS_FINCHOMAGE', StrToDate(DateRepPart));   //Champ détourné
  T.PutValue('PAS_DATEARRET', StrToDate(DateArret));
  T.PutValue('PAS_NONREPRIS', NonRepris);
  T.PutValue('PAS_REPRISEPARTIEL', RepPartiel);
  T.PutValue('PAS_PLANSOC', TravNuit);                    //Champ détourné
  T.PutValue('PAS_RECLASS', ExpoRisk);                    //Champ détourné
  T.PutValue('PAS_DATEPLANSOC', StrToDate(DateSupCtr));   //Champ détourné
  T.PutValue('PAS_CASGEN', Cas);
  T.PutValue('PAS_MONTANT', FloatToStr(Montant));
  T.PutValue('PAS_PLUSDE', Plus);
  T.PutValue('PAS_PERIODEDEBUT', StrToDate(DateDebPer));
  T.PutValue('PAS_PERIODEFIN', StrToDate(DateFinPer));
  T.PutValue('PAS_SUBDEBUT', StrToDate(DateDebSubr));
  T.PutValue('PAS_SUBFIN', StrToDate(DateFinSubr));
  T.PutValue('PAS_TYPMAINTIEN', TMaint);                  //Champ détourné
  T.PutValue('PAS_ORGAUTRE', MoisSubr);                   //Champ détourné
  T.PutValue('PAS_INDPREAVIS', FloatToStr(MntSubro));     //Champ détourné
  T.PutValue('PAS_ORGAGIRC', CKMntsubro);                 //Champ détourné
  T.PutValue('PAS_DECLARLIEU', LieuDecla);
  T.PutValue('PAS_DECLARDATE', StrToDate(DateDecla));
  T.PutValue('PAS_DECLARPERS', PersDecla);
  T.PutValue('PAS_DECLARQUAL', QualDecla);
  T.InsertOrUpdateDB;
  If Tob_Mere <> nil Then Tob_Mere.free;

  //MAJ Tableau
  Tob_Mere := Tob.create('ATTSALAIRES', nil, -1);

  for i := 1 to 12 do
  begin //B For
    DateDebPer := DateToStr(Idate1900);
    IF Trim(Salaires.CellValues[0,i]) <> '' Then DateDebPer := Salaires.CellValues[0,i];
    DateFinPer := DateToStr(Idate1900);
    IF Trim(Salaires.CellValues[1,i]) <> '' Then DateFinPer := Salaires.CellValues[1,i];
    StSal := '0';
    IF Trim(Salaires.CellValues[2,i]) <> '' Then StSal := Salaires.CellValues[2,i];
    MtRegCotPer1 := '0';
    IF Trim(Salaires.CellValues[3,i]) <> '' Then MtRegCotPer1 := Salaires.CellValues[3,i];
    MtRegCotPer2 := '0';
    IF Trim(Salaires.CellValues[4,i]) <> '' Then MtRegCotPer2 := Salaires.CellValues[4,i];
    Motif := '';
    IF Trim(Salaires.CellValues[5,i]) <> '' Then Motif := Salaires.CellValues[5,i];
    NbHeures := '0';
    IF Trim(Salaires.CellValues[6,i]) <> '' Then NbHeures := Salaires.CellValues[6,i];
    NbHeureCompl := '0';
    IF Trim(Salaires.CellValues[7,i]) <> '' Then NbHeureCompl := Salaires.CellValues[7,i];
    MtSalRet := '0';
    IF Trim(Salaires.CellValues[8,i]) <> '' Then MtSalRet := Salaires.CellValues[8,i];

    T := Tob.create('ATTSALAIRES', Tob_Mere, -1);
    If Ordre > 1 then T.SelectDB('"' + IntToStr(Ordre) + '";"' + IntToStr(i) + '"', nil);
    if T = nil then exit;
    T.PutValue('PAL_SALARIE', Matricule);
    T.PutValue('PAL_TYPEATTEST', 'MSA');
    T.PutValue('PAL_ORDRE', ordre);
    T.PutValue('PAL_MOIS', i);
    T.PutValue('PAL_DATEDEBUT', StrToDate(DateDebPer));
    T.PutValue('PAL_DATEFIN', StrToDate(DateFinPer));
    T.PutValue('PAL_SALAIRE', StrToFloat(StSal));
    T.PutValue('PAL_REGCOTPER1', StrToFloat(MtRegCotPer1));
    T.PutValue('PAL_REGCOTPER2', StrToFloat(MtRegCotPer2));
    T.PutValue('PAL_MOTIF', Motif);
    T.PutValue('PAL_NBHEURES', NbHeures);
    T.PutValue('PAL_NBHEURESCOMPL', NbHeureCompl);
    T.PutValue('PAL_SALAIRERET', StrToFloat(MtSalRet));
  End; //End For
  j := 12;
  for i := 1 to 3 do
  begin
    DateDebPer := DateToStr(Idate1900);
    Edit := THEdit(GetControl('DEBPERTHE' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DateDebPer := Edit.text;
    DateFinPer := DateToStr(Idate1900);
    Edit := THEdit(GetControl('FINPERTHE' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DateFinPer := Edit.text;
    MtRegCotPer1 := '0';
    NumEdit := THNumEdit(GetControl('MTTHE1' + IntToStr(i)));
    if NumEdit <> nil then MtRegCotPer1 := FloatToStr(NumEdit.Value);
    MtRegCotPer2 := '0';
    NumEdit := THNumEdit(GetControl('MTTHE2' + IntToStr(i)));
    if NumEdit <> nil then MtRegCotPer2 := FloatToStr(NumEdit.Value);
    MoisAdh := '';
    TComBo := ThvalComboBox(GetControl('MOISADH' + IntToStr(i)));
    if TComBo <> nil then if TComBo.value <> '' then MoisAdh := TComBo.value;
    DateAdh := DateToStr(Idate1900);
    Edit := THEdit(GetControl('DATEADH' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DateAdh := Edit.text;
    DebPerAdh := DateToStr(Idate1900);
    Edit := THEdit(GetControl('DEBPERADH' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then DebPerAdh := Edit.text;
    FinPerAdh := DateToStr(Idate1900);
    Edit := THEdit(GetControl('FINPERADH' + IntToStr(i)));
    if Edit <> nil then if IsValidDate(Edit.text) then FinPerAdh := Edit.text;
    MtAdh := '0';
    NumEdit := THNumEdit(GetControl('MTADH' + IntToStr(i)));
    if NumEdit <> nil then MtAdh := FloatToStr(NumEdit.Value);

    j := j + 1;
    T := Tob.create('ATTSALAIRES', Tob_Mere, -1);
    If Ordre > 1 then T.SelectDB('"' + IntToStr(Ordre) + '";"' + IntToStr(j) + '"', nil);
    if T = nil then exit;
    T.PutValue('PAL_SALARIE', Matricule);
    T.PutValue('PAL_TYPEATTEST', 'MSA');
    T.PutValue('PAL_ORDRE', ordre);
    T.PutValue('PAL_MOIS', j);
    T.PutValue('PAL_DATEDEBUT', StrToDate(DateDebPer));
    T.PutValue('PAL_DATEFIN', StrToDate(DateFinPer));
    T.PutValue('PAL_REGCOTPER1', StrToFloat(MtRegCotPer1));
    T.PutValue('PAL_REGCOTPER2', StrToFloat(MtRegCotPer2));
    T.PutValue('PAL_MOTIF', MoisAdh);  //Champ détourné
    T.PutValue('PAL_VERSEMENT', StrToDate(DateAdh));
    T.PutValue('PAL_DEBPERVER', StrToDate(DebPerAdh));
    T.PutValue('PAL_FINPERVER', StrToDate(FinPerAdh));
    T.PutValue('PAL_MONTANTVER', StrToFloat(MtAdh));
  End;
  If Tob_Mere <> nil Then
  Begin
    If Tob_Mere.detail.count > 0 then
       Tob_Mere.InsertOrUpdateDB;
    Tob_Mere.free;
  End;
End;

Function TOF_PGMALADIE_MSA.SaisieNonValide: Boolean;
Begin
  Result := False;
  If (GetControlText('CKMOTIFMAL') = '-')
  AND (GetControlText('CKMOTIFMAT') = '-')
  AND (GetControlText('CKMOTIFPAT') = '-')
  AND (GetControlText('CKMOTIFJOURMAT') = '-') then
    Begin
      result := True;
      PgiBox('Vous devez renseigner le type d''attestation (Page 1).',Ecran.Caption);  //PT2 Ajout (Page 1)
    End;
End;

Procedure TOF_PGMALADIE_MSA.InitAttes;
var
  QAttes : TQuery;
  TComBo : ThvalComboBox;
  Edit   : THEdit;
  NumEdit: THNumEdit;
  TMaint, cas, CKMntsubro, Ordre : string;
  i,j : integer;
  TXT : String;   //PT2
Begin
  Loaded := True;
  QAttes := OpenSql('SELECT PAS_ORDRE,PAS_MOTIFARRET,PAS_DERNIERJOUR,PAS_DEBCHOMAGE,PAS_REPRISEARRET,PAS_SITUATION,' +
     'PAS_CONVNUMERO,PAS_FINCHOMAGE,PAS_DATEARRET,PAS_NONREPRIS,PAS_REPRISEPARTIEL,PAS_PLANSOC,PAS_RECLASS,' +
     'PAS_DATEPLANSOC,PAS_CASGEN,PAS_MONTANT,PAS_PLUSDE,PAS_PERIODEDEBUT,PAS_PERIODEFIN,PAS_SUBDEBUT,PAS_SUBFIN,' +
     'PAS_TYPMAINTIEN,PAS_ORGAUTRE,PAS_INDPREAVIS,PAS_ORGAGIRC,PAS_DECLARLIEU,PAS_DECLARDATE,' +
     'PAS_DECLARPERS,PAS_DECLARQUAL FROM ATTESTATIONS ' +
    'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="MSA" AND PAS_ASSEDICCAISSE="MAL"', True);

  If not QAttes.EOF Then //PORTAGECWAS
  Begin
     Ordre := QAttes.FindField('PAS_ORDRE').asstring;
     If QAttes.FindField('PAS_MOTIFARRET').asstring = 'MAL' then SetControlText('CKMOTIFMAL','X');
     If QAttes.FindField('PAS_MOTIFARRET').asstring = 'MAT' then SetControlText('CKMOTIFMAT','X');
     If QAttes.FindField('PAS_MOTIFARRET').asstring = 'PAT' then SetControlText('CKMOTIFPAT','X');
     If QAttes.FindField('PAS_MOTIFARRET').asstring = 'MJT' then SetControlText('CKMOTIFJOURMAT','X');
     If (QAttes.FindField('PAS_DERNIERJOUR').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATEDERTRAV', DateToStr(QAttes.FindField('PAS_DERNIERJOUR').AsDateTime));
         TXT := FImpr(GetControlText('DATEDERTRAV'),'N');  //PT2
         SetControlText('DATEDERTRAV1', TXT);              //PT2
     End; //PT2
     If (QAttes.FindField('PAS_DEBCHOMAGE').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATEDEBPATER', DateToStr(QAttes.FindField('PAS_DEBCHOMAGE').AsDateTime));
         TXT := FImpr(GetControlText('DATEDEBPATER'),'N');  //PT2
         SetControlText('DATEDEBPATER1', TXT);              //PT2
     End; //PT2
     If (QAttes.FindField('PAS_REPRISEARRET').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATEREPTRAV', DateToStr(QAttes.FindField('PAS_REPRISEARRET').AsDateTime));
         TXT := FImpr(GetControlText('DATEREPTRAV'),'N');  //PT2
         SetControlText('DATEREPTRAV1', TXT);              //PT2
     End; //PT2
     SetControlText('SITUATIONARRET', QAttes.FindField('PAS_SITUATION').asstring);
     SetControlText('NBENFANT', QAttes.FindField('PAS_CONVNUMERO').asstring);
     If (QAttes.FindField('PAS_FINCHOMAGE').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATEREPRPATIEL', DateToStr(QAttes.FindField('PAS_FINCHOMAGE').AsDateTime));
         TXT := FImpr(GetControlText('DATEREPRPATIEL'),'N');  //PT2
         SetControlText('DATEREPRPATIEL1', TXT);              //PT2
     End; //PT2
     If (QAttes.FindField('PAS_DATEARRET').AsDateTime <> IDate1900) then
         SetControlText('DATEARRET', DateToStr(QAttes.FindField('PAS_DATEARRET').AsDateTime));
     SetControlText('CKNONREPRIS',QAttes.FindField('PAS_NONREPRIS').asstring);
     SetControlText('CKMOTIFPER',QAttes.FindField('PAS_REPRISEPARTIEL').asstring);
     SetControlText('CKTRAVNUIT',QAttes.FindField('PAS_PLANSOC').asstring);
     SetControlText('CKEXPORISQ',QAttes.FindField('PAS_RECLASS').asstring);
     If (QAttes.FindField('PAS_DATEPLANSOC').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATESUSPCRT', DateToStr(QAttes.FindField('PAS_DATEPLANSOC').AsDateTime));
         TXT := FImpr(GetControlText('DATESUSPCRT'),'N');  //PT2
         SetControlText('DATESUSPCRT1', TXT);              //PT2
     End; //PT2

     cas := QAttes.FindField('PAS_CASGEN').asstring;
     If cas = 'GEN' then
     Begin
        SetControlText('CKCASGEN','X');
        NumEdit := THNumEdit(GetControl('MTCOT200'));
        If (QAttes.FindField('PAS_MONTANT').asFloat <> 0) And (NumEdit <> nil) Then
           NumEdit.Value := QAttes.FindField('PAS_MONTANT').asFloat;
        SetControlText('CK200',QAttes.FindField('PAS_PLUSDE').asstring);
     End
     Else If cas = 'PAR' then
     Begin
        SetControlText('CKCASPART','X');
        NumEdit := THNumEdit(GetControl('MTCOT800'));
        If (QAttes.FindField('PAS_MONTANT').asFloat <> 0) And (NumEdit <> nil) Then
           NumEdit.Value := QAttes.FindField('PAS_MONTANT').asFloat;
        SetControlText('CK800',QAttes.FindField('PAS_PLUSDE').asstring);
     End;
     If (QAttes.FindField('PAS_PERIODEDEBUT').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATEDEBPER', DateToStr(QAttes.FindField('PAS_PERIODEDEBUT').AsDateTime));
         TXT := FImpr(GetControlText('DATEDEBPER'),'C');  //PT2
         SetControlText('DATEDEBPER1', TXT);              //PT2
     End; //PT2
     If (QAttes.FindField('PAS_PERIODEFIN').AsDateTime <> IDate1900) then
     Begin //PT2
         SetControlText('DATEFINPER', DateToStr(QAttes.FindField('PAS_PERIODEFIN').AsDateTime));
         TXT := FImpr(GetControlText('DATEFINPER'),'C');  //PT2
         SetControlText('DATEFINPER1', TXT);              //PT2
     End; //PT2
     If (QAttes.FindField('PAS_SUBDEBUT').AsDateTime <> IDate1900) then
         SetControlText('DATEDEBSUBR', DateToStr(QAttes.FindField('PAS_SUBDEBUT').AsDateTime));
     If (QAttes.FindField('PAS_SUBFIN').AsDateTime <> IDate1900) then
         SetControlText('DATEFINSUBR', DateToStr(QAttes.FindField('PAS_SUBFIN').AsDateTime));

     TMaint := QAttes.FindField('PAS_TYPMAINTIEN').asstring;
     If TMaint = 'INT' Then SetControlText('CKSUBINT','X');
     If TMaint = 'PAR' Then SetControlText('CKSUBPART','X');
     SetControlText('SUBMOIS',QAttes.FindField('PAS_ORGAUTRE').asstring);
     If (QAttes.FindField('PAS_INDPREAVIS').asFloat <> 0) Then
     Begin
        NumEdit := THNumEdit(GetControl('SUBMT'));
        If NumEdit <> nil Then NumEdit.Value := QAttes.FindField('PAS_INDPREAVIS').asFloat;
     End;
     CKMntsubro := QAttes.FindField('PAS_ORGAGIRC').asstring;
     If CKMntsubro = 'X' Then SetControlText('CKSUBOUI','X');
     If CKMntsubro = '-' Then SetControlText('CKSUBNON','X');
     SetControlText('LIEUFAIT',QAttes.FindField('PAS_DECLARLIEU').asstring);
     If (QAttes.FindField('PAS_DECLARDATE').AsDateTime <> IDate1900) then
         SetControlText('DATEFAIT', DateToStr(QAttes.FindField('PAS_DECLARDATE').AsDateTime));
     SetControlText('NOMSIGNE',QAttes.FindField('PAS_DECLARPERS').asstring);
     SetControlText('QUALITESIGNE',QAttes.FindField('PAS_DECLARQUAL').asstring);
  End;
  Ferme(QAttes);
  //Init Tableau
  QAttes := OpenSql('SELECT PAL_SALARIE,PAL_TYPEATTEST,PAL_ORDRE,PAL_MOIS,PAL_DATEDEBUT,' +
       'PAL_DATEFIN,PAL_SALAIRE,PAL_REGCOTPER1,PAL_REGCOTPER2,PAL_MOTIF,PAL_NBHEURES,PAL_NBHEURESCOMPL,' +
       'PAL_SALAIRERET,PAL_VERSEMENT,PAL_DEBPERVER,PAL_FINPERVER,PAL_MONTANTVER ' +
       'FROM ATTSALAIRES WHERE PAL_SALARIE="' + Matricule + '" AND PAL_TYPEATTEST="MSA" ' +
       'And PAL_ORDRE = '+Ordre+' ORDER BY PAL_MOIS', True);

  i := 0;
  j := 0;
  While not QAttes.EOF Do
  Begin //B While
    i := i + 1;
    If i <= 12 Then
    Begin
       If (QAttes.FindField('PAL_DATEDEBUT').AsDateTime <> IDate1900) then
           Salaires.CellValues[0,i] := DateToStr(QAttes.FindField('PAL_DATEDEBUT').AsDateTime);
       If (QAttes.FindField('PAL_DATEFIN').AsDateTime <> IDate1900) then
           Salaires.CellValues[1,i] := DateToStr(QAttes.FindField('PAL_DATEFIN').AsDateTime);
       If (QAttes.FindField('PAL_SALAIRE').asFloat <> 0) Then
           Salaires.CellValues[2,i] := FormatFloat('#0.00',QAttes.FindField('PAL_SALAIRE').asFloat);    //PT2 '# ##0.00'
       If (QAttes.FindField('PAL_REGCOTPER1').asFloat <> 0) Then
           Salaires.CellValues[3,i] := FormatFloat('#0.00',QAttes.FindField('PAL_REGCOTPER1').asFloat); //PT2 '# ##0.00'
       If (QAttes.FindField('PAL_REGCOTPER2').asFloat <> 0) Then
           Salaires.CellValues[4,i] := FormatFloat('#0.00',QAttes.FindField('PAL_REGCOTPER2').asFloat); //PT2 '# ##0.00'
       If (QAttes.FindField('PAL_MOTIF').asstring <> '') Then
           Salaires.CellValues[5,i] := QAttes.FindField('PAL_MOTIF').asstring;
       If (QAttes.FindField('PAL_NBHEURES').asstring <> '0') Then
           Salaires.CellValues[6,i] := QAttes.FindField('PAL_NBHEURES').asstring;
       If (QAttes.FindField('PAL_NBHEURESCOMPL').asstring <> '0') Then
           Salaires.CellValues[7,i] := QAttes.FindField('PAL_NBHEURESCOMPL').asstring;
       If (QAttes.FindField('PAL_SALAIRERET').asFloat <> 0) Then
           Salaires.CellValues[8,i] := FormatFloat('#0.00',QAttes.FindField('PAL_SALAIRERET').asFloat); //PT2 '# ##0.00'
    End
    Else
    Begin
       j := j + 1;
       Edit := THEdit(GetControl('DEBPERTHE' + IntToStr(j)));
       If (Edit <> nil) and (QAttes.FindField('PAL_DATEDEBUT').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_DATEDEBUT').AsDateTime);
       Edit := THEdit(GetControl('FINPERTHE' + IntToStr(j)));
       If (Edit <> nil) and (QAttes.FindField('PAL_DATEFIN').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_DATEFIN').AsDateTime);
       NumEdit := THNumEdit(GetControl('MTTHE1' + IntToStr(j)));
       If (NumEdit <> nil) And (QAttes.FindField('PAL_REGCOTPER1').asFloat <> 0) Then
           NumEdit.Value := QAttes.FindField('PAL_REGCOTPER1').asFloat;
       NumEdit := THNumEdit(GetControl('MTTHE2' + IntToStr(j)));
       If (NumEdit <> nil) And (QAttes.FindField('PAL_REGCOTPER2').asFloat <> 0) Then
           NumEdit.Value := QAttes.FindField('PAL_REGCOTPER2').asFloat;
       TComBo := ThvalComboBox(GetControl('MOISADH' + IntToStr(j)));
       If (TComBo <> nil) And (QAttes.FindField('PAL_MOTIF').asstring <> '') Then
           TComBo.value := QAttes.FindField('PAL_MOTIF').asstring;
       Edit := THEdit(GetControl('DATEADH' + IntToStr(j)));
       If (Edit <> nil) and (QAttes.FindField('PAL_VERSEMENT').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_VERSEMENT').AsDateTime);
       Edit := THEdit(GetControl('DEBPERADH' + IntToStr(j)));
       If (Edit <> nil) and (QAttes.FindField('PAL_DEBPERVER').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_DEBPERVER').AsDateTime);
       Edit := THEdit(GetControl('FINPERADH' + IntToStr(j)));
       If (Edit <> nil) and (QAttes.FindField('PAL_FINPERVER').AsDateTime <> IDate1900) then
           Edit.Text := DateToStr(QAttes.FindField('PAL_FINPERVER').AsDateTime);
       NumEdit := THNumEdit(GetControl('MTADH' + IntToStr(j)));
       If (NumEdit <> nil) And (QAttes.FindField('PAL_MONTANTVER').asFloat <> 0) Then
           NumEdit.Value := QAttes.FindField('PAL_MONTANTVER').asFloat;
    End;
    QAttes.next;
  End; //E While
  Ferme(QAttes);
  Loaded := False;
End;

Procedure TOF_PGMALADIE_MSA.InitPages(Var Complet : Boolean);
Var
  i : Integer;
Begin
  If Complet = TRUE Then
  Begin
     SetControlText('CKMOTIFMAL','-');
     SetControlText('CKMOTIFMAT','-');
     SetControlText('CKMOTIFPAT','-');
     SetControlText('CKMOTIFJOURMAT','-');
     SetControlText('DATEDERTRAV','01/01/1900');   //PT2
     SetControlText('DATEDERTRAV1','');            //PT2
     SetControlText('DATEDEBPATER','01/01/1900');  //PT2
     SetControlText('DATEDEBPATER1','');           //PT2
     SetControlText('DATEREPTRAV','01/01/1900');   //PT2
     SetControlText('DATEREPTRAV1','');            //PT2
     SetControlText('SITUATIONARRET','');
     SetControlText('NBENFANT','');
     SetControlText('DATEREPRPATIEL','01/01/1900');//PT2
     SetControlText('DATEREPRPATIEL1','');         //PT2
     SetControlText('DATEARRET','01/01/1900');     //PT2
     SetControlText('CKNONREPRIS','-');
     SetControlText('CKMOTIFPER','-');
     SetControlText('CKTRAVNUIT','-');
     SetControlText('CKEXPORISQ','-');
     SetControlText('DATESUSPCRT','01/01/1900');   //PT2
     SetControlText('DATESUSPCRT1','');            //PT2
     SetControlText('CKCASGEN','-');
     SetControlText('CKCASPART','-');
     SetControlText('DATEDEBSUBR','01/01/1900');   //PT2
     SetControlText('DATEFINSUBR','01/01/1900');   //PT2
     SetControlText('CKSUBINT','-');
     SetControlText('CKSUBPART','-');
     SetControlText('SUBMOIS','');
     SetControlText('SUBMT','');
     SetControlText('CKSUBOUI','-');
     SetControlText('CKSUBNON','-');
     SetControlText('LIEUFAIT','');
     SetControlText('DATEFAIT','01/01/1900');      //PT2
     SetControlText('NOMSIGNE','');
     SetControlText('QUALITESIGNE','');
     For i := 1 to 3 do
     Begin
        SetControlText('DEBPERTHE' + IntToStr(i),'01/01/1900');   //PT2
        SetControlText('FINPERTHE' + IntToStr(i),'01/01/1900');   //PT2
        SetControlText('MTTHE1' + IntToStr(i),'');
        SetControlText('MTTHE2' + IntToStr(i),'');
        SetControlText('MOISADH' + IntToStr(i),'');
        SetControlText('DATEADH' + IntToStr(i),'01/01/1900');     //PT2
        SetControlText('DEBPERADH' + IntToStr(i),'01/01/1900');   //PT2
        SetControlText('FINPERADH' + IntToStr(i),'01/01/1900');   //PT2
        SetControlText('MTADH' + IntToStr(i),'');
     End;
  End;
  SetControlText('MTCOT200','');
  SetControlText('CK200','-');
  SetControlText('MTCOT800','');
  SetControlText('CK800','-');
  SetControlText('DATEDEBPER','01/01/1900');  //PT2
  SetControlText('DATEDEBPER1','');           //PT2
  SetControlText('DATEFINPER','01/01/1900');  //PT2
  SetControlText('DATEFINPER1','');           //PT2
  For i := 1 to 12 do
  Begin
     Salaires.CellValues[0,i] := '';
     Salaires.CellValues[1,i] := '';
     Salaires.CellValues[2,i] := '';
     Salaires.CellValues[3,i] := '';
     Salaires.CellValues[4,i] := '';
     Salaires.CellValues[5,i] := '';
     Salaires.CellValues[6,i] := '';
     Salaires.CellValues[7,i] := '';
     Salaires.CellValues[8,i] := '';
  End;
End;

Procedure TOF_PGMALADIE_MSA.Imprimer(Sender: TObject);
var
  Pages : TPageControl;
  StPages, Requete : String;
Begin
  if SaisieNonValide then Exit;
  Pages := TPageControl(GetControl('PAGES'));
  Requete := '';
  StPages := AglGetCriteres (Pages, FALSE);
  LanceEtat('E','PAT','ML1',True,False,False,Pages,Requete,'',False,0,StPages);
  LanceEtat('E','PAT','ML2',True,False,False,Pages,Requete,'',False,0,StPages);
End;

Procedure TOF_PGMALADIE_MSA.OnLoad;
Var
  Page     : TPageControl;
Begin
  Loaded := False;
  Page := TPageControl(GetControl('PAGES'));
  If Page <> nil Then Page.ActivePageIndex := 0;
End;

Procedure TOF_PGMALADIE_MSA.OnNew;
Var
  AFaire : Boolean;
Begin
  AFaire := True;
  InitPages(AFaire);
End;

Procedure TOF_PGMALADIE_MSA.DateElipsisclick(Sender: TObject);
var
  key: char;
Begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
End;

initialization
  registerclasses([TOF_PGMALADIE_MSA]);
end.

