{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 27/12/2006
Modifié le ... :   /  /
Description .. : Edition de la liste du personnel entre 2 dates
Mots clefs ... : PAIE;PERSONNEL
*****************************************************************}
{
PT1 15/01/07 V8_00 FCO Mise en place filtrage des habilitations/poupulations
PT2 18/04/07 V_72  FC Suppression du LanceEtatTob
PT3 09/07/07 V_80  FC FQ 12177
}
unit UTOFPGListePersonnelDuAu;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UtilEAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1, EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc, HQry, UTob, HTB97, ed_tools,
  HStatus;



type
  TOF_PGListePersonnelDuAu = class(TOF)
  private
    TobEtat, T: Tob;
    procedure Change(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ConstruireTob();
    procedure RemplirInfosTob(TR: Tob);
    Function RendLibQualif(Qualif,Convention : String) : String;
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override; //PT2
    procedure OnClose; override ; //PT2


  end;
implementation

uses PgEditOutils, PGEditOutils2,EntPaie, PgOutils, PGEdtEtat,
  P5Def, StrUtils;  // VisibiliteChampSalarie et VisibiliteChampLibreSal

procedure TOF_PGListePersonnelDuAu.OnArgument(Arguments: string);
var
  Defaut: THEdit;
  CEtab: TCheckBox;
  XRupture, XOrder: THEdit;
  Min, Max: string;
  Ch1, Ch2, Ch3: string;
  Num: Integer;
begin
  inherited;
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    Defaut.text := GetParamSoc('SO_LIBELLE');

  for Num := 1 to 4 do
  begin
    Ch1 := 'PSA_TRAVAILN' + IntToStr(Num);
    Ch2 := 'TPSA_TRAVAILN' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_TRAVAILN' + IntToStr(Num);
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;

  for Num := 1 to 4 do
  begin
    Ch1 := 'PSA_LIBREPCMB' + IntToStr(Num);
    Ch2 := 'TPSA_LIBREPCMB' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_LIBREPCMB' + IntToStr(Num);
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;

  //Valeur par défaut
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  Defaut := ThEdit(getcontrol('PSA_DATEENTREE'));
  if Defaut <> nil then Defaut.text := DatetoStr(Date);
  Defaut := ThEdit(getcontrol('PSA_DATEENTREE_'));
  if Defaut <> nil then Defaut.text := DatetoStr(Date);

  CEtab := TCheckBox(GetControl('CETAB'));
  if CEtab <> nil then
  begin
    CEtab.OnClick := Change;
    Cetab.Checked := True;
  end;

  Defaut := ThEdit(getcontrol('PSA_DATEENTREE'));
  if Defaut <> nil then
  begin
    Defaut.OnElipsisClick := DateElipsisclick;
    Defaut.OnExit := Change;
  end;
  Defaut := ThEdit(getcontrol('PSA_DATEENTREE_'));
  if Defaut <> nil then
  begin
    Defaut.OnElipsisClick := DateElipsisclick;
    Defaut.OnExit := Change;
  end;

  XRupture := THEdit(GetControl('XX_RUTURE1'));
  XOrder := THEdit(GetControl('XX_ORDERBY'));
  if (Xorder <> nil) then Xorder.text :=   ' PSA_DATEENTREE,PSA_SALARIE,PCI_TYPECONTRAT DESC';
  if (XRupture <> nil) then XRupture.text := 'PSA_ETABLISSEMENT';
end;

procedure TOF_PGListePersonnelDuAu.Change(Sender: TObject);
var
  XOrder, XRupture: THEDIT;
  CEtab: TCheckBox;
begin
  XRupture := THEdit(GetControl('XX_RUPTURE1'));
  CEtab := TCheckBox(GetControl('CETAB'));
  XOrder := THEdit(GetControl('XX_ORDERBY'));
  if (Xorder <> nil) and (cetab <> nil) and (XRupture <> nil) then
  begin
    if cetab.Checked = true then
    begin
      Xorder.text :=  ' PSA_DATEENTREE,PSA_SALARIE,PCI_TYPECONTRAT DESC';
      XRupture.text := 'PSA_ETABLISSEMENT';
    end;
    if cetab.Checked = False then
    begin
      Xorder.text := ' PSA_DATEENTREE,PSA_SALARIE,PCI_TYPECONTRAT DESC';
      XRupture.text := '';
    end;
  end;
  if not IsValidDate(GetControlText('PSA_DATEENTREE')) then
  begin
    PGIBox('La date de présence "du" est erronée.', Ecran.caption);
    SetControlText('PSA_DATEENTREE', DatetoStr(Date));
  end;
  if not IsValidDate(GetControlText('PSA_DATEENTREE_')) then
  begin
    PGIBox('La date de présence "au" est erronée.', Ecran.caption);
    SetControlText('PSA_DATEENTREE_', DatetoStr(Date));
  end;
end;

procedure TOF_PGListePersonnelDuAu.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGListePersonnelDuAu.ConstruireTob();
begin
  T := Tob.Create('Les filles salariés', TobEtat, -1);
  T.AddChampSup('PSA_SALARIE', False);
  T.AddChampSup('PSA_LIBELLE', False);
  T.AddChampSup('PSA_PRENOM', False);
  T.AddChampSup('PSA_ETABLISSEMENT', False);
  T.AddChampSup('PSA_ADRESSE1', False);
  T.AddChampSup('PSA_ADRESSE2', False);
  T.AddChampSup('PSA_ADRESSE3', False);
  T.AddChampSup('PSA_VILLE', False);
  T.AddChampSup('PSA_CODEPOSTAL', False);
  T.AddChampSup('PSA_DATENAISSANCE', False);
  T.AddChampSup('PSA_SEXE', False);
  T.AddChampSup('PSA_NUMEROSS', False);
  T.AddChampSup('PSA_NATIONALITE', False);
  T.AddChampSup('PSA_LIBELLEEMPLOI', False);
  T.AddChampSup('PSA_QUALIFICATION', False);
  T.AddChampSup('PSA_DATEENTREE', False);
  T.AddChampSup('PSA_DATESORTIE', False);
  T.AddChampSup('PSA_CONDEMPLOI', False);
  T.AddChampSup('PSA_CARTESEJOUR', False);
  T.AddChampSup('PSA_DELIVPAR', False);
  T.AddChampSup('SEJOUR', False);
  T.AddChampSup('PCI_TYPECONTRAT', False);
  T.AddChampSup('AGENCE', False);
  T.AddChampSup('ADRESSE1', False);
  T.AddChampSup('ADRESSE2', False);
  T.AddChampSup('CPVILLE', False);
  T.AddChampSup('PSA_COEFFICIENT', False);
  T.AddChampSup('TYPEPERSONNE', False);
  T.AddChampSup('MUTATION', False);
  T.AddChampSup('PSA_MOTIFSORTIE', False);
  T.AddChampSup('DATELICENCIEMENT', False);
  T.AddChampSup('PSA_CATDADS', False);
  T.AddChampSup('PSA_DADSCAT', False);
  T.AddChampSup('PSA_INDICE', False);
  T.AddChampSup('PSA_CONVENTION', False);
  T.AddChampSup('LIBQUALIF', False);
  T.AddChampSup('DUREECONTRAT', False);
end;

procedure TOF_PGListePersonnelDuAu.RemplirInfosTob(TR: Tob);
var LaQualif,LaConvention : String;
begin
  T.PutValue('PSA_SALARIE', TR.GetValue('PSA_SALARIE'));
  T.PutValue('PSA_LIBELLE', TR.GetValue('PSA_LIBELLE'));
  T.PutValue('PSA_PRENOM', TR.GetValue('PSA_PRENOM'));
  T.PutValue('PSA_ADRESSE1', TR.GetValue('PSA_ADRESSE1'));
  T.PutValue('PSA_ADRESSE2', TR.GetValue('PSA_ADRESSE2'));
  T.PutValue('PSA_ADRESSE3', TR.GetValue('PSA_ADRESSE3'));
  T.PutValue('PSA_VILLE', TR.GetValue('PSA_VILLE'));
  T.PutValue('PSA_CODEPOSTAL', TR.GetValue('PSA_CODEPOSTAL'));
  T.PutValue('PSA_DATENAISSANCE', TR.GetValue('PSA_DATENAISSANCE'));
  T.PutValue('PSA_SEXE', TR.GetValue('PSA_SEXE'));
  T.PutValue('PSA_NUMEROSS', TR.GetValue('PSA_NUMEROSS'));
  T.PutValue('PSA_NATIONALITE', TR.GetValue('PSA_NATIONALITE'));
  T.PutValue('AGENCE', '');
  T.PutValue('ADRESSE1', '');
  T.PutValue('DATELICENCIEMENT', IDate1900);
  T.PutValue('ADRESSE2', '');
  T.PutValue('CPVILLE', '');
  T.PutValue('PSA_CONDEMPLOI', TR.GetValue('PSA_CONDEMPLOI'));
  T.PutValue('PSA_CARTESEJOUR', TR.GetValue('PSA_CARTESEJOUR'));
  T.PutValue('PSA_DELIVPAR', TR.GetValue('PSA_DELIVPAR'));
  T.PutValue('SEJOUR', TR.GetValue('PSA_DATEXPIRSEJOUR'));
  T.PutValue('PSA_COEFFICIENT', TR.getValue('PSA_COEFFICIENT'));
  T.PutValue('PSA_CATDADS', TR.getValue('PSA_CATDADS'));
  T.PutValue('PSA_DADSCAT', TR.getValue('PSA_DADSCAT'));
  T.PutValue('PSA_INDICE', TR.getValue('PSA_INDICE'));
  If TR.getValue('PSA_CONVENTION') <> null then LaConvention := TR.GetValue('PSA_CONVENTION')
  else LaConvention := '';
  If TR.getValue('PSA_QUALIFICATION') <> null then LaQualif := TR.GetValue('PSA_QUALIFICATION')
  else LaQualif := '';
  T.PutValue('PSA_CONVENTION', LaConvention);
  T.PutValue('PSA_QUALIFICATION', LaQualif);
  T.PutValue('PSA_LIBELLEEMPLOI', TR.getValue('PSA_LIBELLEEMPLOI'));
  T.PutValue('LIBQUALIF','');
end;

function TOF_PGListePersonnelDuAu.RendLibQualif(Qualif,Convention : String) : String;
var Q : TQuery;
begin
        If qualif = '' then
        begin
             Result := '';
             Exit;
        end;
        Q := OpenSQL('SELECT PMI_LIBELLE FROM MINIMUMCONVENT '+
        'WHERE (PMI_CONVENTION="'+Convention+'" OR PMI_CONVENTION="000") AND ##PMI_PREDEFINI## PMI_CODE="'+Qualif+'" AND PMI_NATURE="QUA"',True); //PT19
        If Not Q.eof then Result := Q.FindField('PMI_LIBELLE').AsString
        else Result := '';
        Ferme(Q);
end;

procedure TOF_PGListePersonnelDuAu.OnClose;
begin
  inherited;
  FreeAndNil (TobEtat); //PT2
end;

procedure TOF_PGListePersonnelDuAu.OnUpdate;
var
  Q, QContrat: TQuery;
  TobPaie, TobSalarie, TobInterim, TobEtab, TC, TobAgence, TA, TobLesContrats: Tob;
  a, i, p, j, x, c: Integer;
  Trouver, ChangementPaie, ChangementEtab, TrouveSortie: Boolean;
  Pages: TPageControl;
  DSalInt, DSalarie, DateEdition, DateLegale: TdateTime;
  Salarie, Etablissement, NouvelEtab, StPages, LibelleEmploi: string;
  Qualification,Convention,LibQualif : String;
  DateSortie, DateEntree, DateContrat, DateChangeEtab, AncDateSortie, NouvDateEntree, DateFCt, DateDCt: TDateTime;
  Etab1,Etab2 : String;
  Travail : String;
  LibRepCmb : String;
  Requete : String;
  Num : Integer;
  T1  : String;
  T2  : String;
  DateDeb, Date1, Date2: TDateTime;
  PremMois, PremAnnee, duree: Word;
begin
  inherited;

  FreeAndNil(TobEtat);

  DateEdition := StrToDate(GetControlText('PSA_DATEENTREE'));
  DateLegale := StrToDate(GetControlText('PSA_DATEENTREE_'));
  Pages := TPageControl(GetControl('Pages'));
  Etab1 := GetcontrolText('PSA_ETABLISSEMENT');
  Etab2 := GetcontrolText('PSA_ETABLISSEMENT_');
  Q := OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_ETABLISSEMENT>="'+Etab1+'" '+
  'AND ET_ETABLISSEMENT<="'+Etab2+'"',True);
  TobEtab:= Tob.Create('Etab',Nil,-1);
  TobEtab.LoadDetailDB('Etab','','',Q,False);
  Ferme(Q);
  Q := OpenSQL('SELECT ANN_GUIDPER,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE FROM ANNUAIRE WHERE ANN_TYPEPER="AGI"', True);
  TobAgence := Tob.create('Les agences', nil, -1);
  TobAgence.LoadDetailDB('ANNUAIRE', '', '', Q, False);
  Ferme(Q);
  TobEtat := Tob.Create('Les salariés pour édition', nil, -1);

  Requete := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,' +
      'PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,PSA_VILLE,PSA_CODEPOSTAL,' +
      'PSA_DATENAISSANCE,PSA_SEXE,PSA_NUMEROSS,PSA_NATIONALITE,PSA_LIBELLEEMPLOI,PSA_CONVENTION,PSA_QUALIFICATION,PSA_COEFFICIENT,' +
      'PSA_DATEENTREE,PSA_DATESORTIE,PSA_CATDADS,PSA_DADSCAT,PSA_MOTIFSORTIE,' +
      'PSA_CARTESEJOUR,PSA_DELIVPAR,PSA_DATEXPIRSEJOUR,PSA_CONDEMPLOI,PSA_INDICE' +
      ' FROM SALARIES '+
      'WHERE '+ SqlConf ('SALARIES') +
      'AND PSA_DATEENTREE<="' + UsDateTime(DateLegale) + '" AND ' +
      'AND (PSA_DATESORTIE>="' + UsDateTime(DateEdition) + '" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND ' +
      '((PSA_ETABLISSEMENT>="' + Etab1 + '" AND PSA_ETABLISSEMENT<="'+Etab2+'")' +
      ' OR (PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEDEBUT<="' + UsDateTime(DateLegale) + '"' +
      ' AND PPU_DATEFIN>="' + UsDateTime(DateEdition) + '" AND PPU_ETABLISSEMENT>="'+ Etab1+'" AND PPU_ETABLISSEMENT<="'+etab2+'"))'+
      ' OR (PSA_SALARIE IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE PCI_DEBUTCONTRAT<="' + UsDateTime(DateLegale) + '"' +
      ' AND (PCI_FINCONTRAT>="' + UsDateTime(DateEdition) + '" OR PCI_FINCONTRAT<="'+UsDateTime(IDate1900)+'") AND PCI_ETABLISSEMENT>="'+ Etab1+'" AND PCI_ETABLISSEMENT<="'+etab2+'")))';

  // Récupérer les valeurs des zones libres et les inclure dans la clause WHERE
  for Num := 1 to 4 do
  begin
    Travail := GetControlText('PSA_TRAVAILN' + IntToStr(Num));
    if (Travail <> '') and (Travail <> '<<Tous>>') then
    begin
      T1 := READTOKENST(Travail);
      T2 := '';
      while T1 <> '' do
      begin
        if T2 <> '' then
          T2 := T2 + ',';
        T2 := T2 + '"' + T1 + '"';
        T1 := READTOKENST(Travail);
      end;
      Requete := Requete + ' and PSA_TRAVAILN' + IntToStr(Num) + ' IN (' + T2 +')';
    end;
  end;

  for Num := 1 to 4 do
  begin
    Travail := GetControlText('PSA_LIBREPCMB' + IntToStr(Num));
    if (Travail <> '') and (Travail <> '<<Tous>>') then
    begin
      T1 := READTOKENST(Travail);
      T2 := '';
      while T1 <> '' do
      begin
        if T2 <> '' then
          T2 := T2 + ',';
        T2 := T2 + '"' + T1 + '"';
        T1 := READTOKENST(Travail);
      end;
      Requete := Requete + ' and PSA_LIBREPCMB' + IntToStr(Num) + ' IN (' + T2 +')';
    end;
  end;

  //DEB PT1
  if Assigned(MonHabilitation) then
    Requete := Requete + ' AND ' + MonHabilitation.LeSQL;
  //FIN PT1

  Requete := Requete + ' ORDER BY PSA_DATEENTREE';


  Q := OpenSQL(Requete, True);
   TobSalarie := Tob.create('Les salaries', nil, -1);
   TobSalarie.LoadDetailDB('SALARIES', '', '', Q, false);
   Ferme(Q);
   InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', TobSalarie.Detail.Count, False, True);
   InitMove(TobSalarie.Detail.Count, '');
   For i := 0 to TobSalarie.Detail.Count - 1 do
   begin
        Salarie := TobSalarie.Detail[i].GetValue('PSA_SALARIE');
        //Recherche des contrats de tarvail qui vont constituer les périodes
        QContrat := OpenSQL('SELECT PCI_TYPECONTRAT,PCI_LIBELLEEMPLOI,PCI_ETABLISSEMENT,PCI_FINCONTRAT,PCI_DEBUTCONTRAT,PCI_MOTIFSORTIE FROM CONTRATTRAVAIL ' + //PT3
        'WHERE PCI_SALARIE="' + Salarie + '" ' +
            'AND (PCI_FINCONTRAT>="' + UsDateTime(DateEdition) + '" OR PCI_FINCONTRAT<="' + UsDateTime(IDate1900) + '") ' +
            'ORDER BY PCI_DEBUTCONTRAT', True);
        TobLesContrats := Tob.Create('Les contrats', nil, -1);
        TobLesContrats.LoadDetailDB('Les contrats', '', '', QContrat, False);
        Ferme(QContrat);
        If TobLesContrats.Detail.Count > 0 then
        begin
          For c := 0 to TobLesContrats.Detail.Count -1 do
          begin
                Etablissement := TobLesContrats.Detail[c].GetValue('PCI_ETABLISSEMENT');
                If TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etablissement],False) <> nil then
                begin
                  ConstruireTob;
                  if TobLesContrats.Detail[c].GetValue('PCI_LIBELLEEMPLOI')<>null then LibelleEmploi := TobLesContrats.Detail[c].GetValue('PCI_LIBELLEEMPLOI');
                  If LibelleEmploi = '' then LibelleEmploi := TobSalarie.Detail[i].getValue('PSA_LIBELLEEMPLOI'); //PT18
                  T.PutValue('TYPEPERSONNE', 'Salarié');
                  RemplirInfosTob(TobSalarie.Detail[i]);
                  T.PutValue('PSA_ETABLISSEMENT', Etablissement);
                  T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                  T.PutValue('PSA_DATEENTREE', TobLesContrats.Detail[c].GetValue('PCI_DEBUTCONTRAT'));
                  Date1 := TobLesContrats.Detail[c].GetValue('PCI_DEBUTCONTRAT');
                  If TobLesContrats.Detail[c].GetValue('PCI_FINCONTRAT') <> IDate1900 then
                  begin
                    T.PutValue('PSA_DATESORTIE', TobLesContrats.Detail[c].GetValue('PCI_FINCONTRAT'));
                    Date2 := TobLesContrats.Detail[c].GetValue('PCI_FINCONTRAT');
                  end
                  else
                  begin
                    T.PutValue('PSA_DATESORTIE', TobSalarie.Detail[i].getValue('PSA_DATESORTIE'));
                    Date2 := TobSalarie.Detail[i].GetValue('PSA_DATESORTIE');
                  end;
                  //DEB PT3
                  If TobLesContrats.Detail[c].GetValue('PCI_MOTIFSORTIE') <> '' then
                    T.PutValue('PSA_MOTIFSORTIE', TobLesContrats.Detail[c].GetValue('PCI_MOTIFSORTIE'))
                  else
                    T.PutValue('PSA_MOTIFSORTIE', TobSalarie.Detail[i].getValue('PSA_MOTIFSORTIE'));
                  //FIN PT3
                  T.PutValue('MUTATION', 'Contrat de travail');
                  T.PutValue('PCI_TYPECONTRAT', TobLesContrats.Detail[c].GetValue('PCI_TYPECONTRAT'));
                  //DEB PT3
                  if (Date2 <= idate1900) then Date2 := Date;
                  Date2 := Date2 + 1;
                  Duree := 0;
                  if (Date1 <> idate1900) then
                    AglNombreDeMoisComplet(Date1, Date2, PremMois, PremAnnee, Duree);
                  T.PutValue('DUREECONTRAT', IntToStr(Duree) + ' mois');
                  //FIN PT3
                end;
          end;
          FreeAndNil(TobLesContrats);
        end
        else
        begin
          FreeAndNil(TobLesContrats);
          //Si paie de contrat, recherche des différentes dates entrées et sorties, ainsi que changement établissement dans les paies
          Q := OpenSQL('SELECT PPU_DATEENTREE,PPU_DATESORTIE,PPU_LIBELLEEMPLOI,PPU_ETABLISSEMENT,PPU_DATEDEBUT,PPU_DATEFIN,PPU_CONVENTION,PPU_QUALIFICATION '+
          'FROM PAIEENCOURS WHERE PPU_SALARIE="'+Salarie+'" AND PPU_DATEDEBUT<="' + UsDateTime(DateLegale) + '"' +
          ' AND PPU_DATEFIN>="' + UsDateTime(DateEdition) + '" ORDER BY PPU_DATEDEBUT',true);
          TobPaie := Tob.Create('Lespaies',Nil,-1);
          TobPaie.LoadDetailDB('LesPaies','','',Q,False);
          Ferme(Q);
          If TobPaie.Detail.Count > 0 then
          begin
             For p := 0 to TobPaie.detail.Count - 1 do
             begin
                If p = 0 then
                begin
                  DateEntree := TobPaie.Detail[p].GetValue('PPU_DATEENTREE');
                  DateSortie := TobPaie.Detail[p].GetValue('PPU_DATESORTIE');
                  Etablissement := TobPaie.Detail[p].GetValue('PPU_ETABLISSEMENT');
                  If TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etablissement],False) <> nil then
                  begin
                    ConstruireTob;
                    T.PutValue('TYPEPERSONNE', 'Salarié');
                    RemplirInfosTob(TobSalarie.Detail[i]);
                    T.PutValue('PSA_ETABLISSEMENT', Etablissement);
                    if TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI') <> null then
                    begin
                         LibelleEmploi := TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI');
                         T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                    end;
                    if (TobPaie.Detail[p].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p].GetValue('PPU_QUALIFICATION') <> null) then  //PT16
                    begin
                         Qualification := TobPaie.Detail[p].GetValue('PPU_QUALIFICATION');
                         Convention := TobPaie.Detail[p].GetValue('PPU_CONVENTION');
                         If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                         If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                    end;
                    T.PutValue('PSA_DATEENTREE', TobPaie.Detail[p].GetValue('PPU_DATEENTREE'));
                    T.PutValue('PSA_DATESORTIE', TobPaie.Detail[p].GetValue('PPU_DATESORTIE'));
                    T.PutValue('PSA_MOTIFSORTIE', '');
                    T.PutValue('MUTATION', '');
                    T.PutValue('PCI_TYPECONTRAT', 'CDI');
                  end;
                end;
                If TobPaie.Detail[p].GetValue('PPU_DATEENTREE') > DateEntree then
                begin
                        T := TobEtat.FindFirst(['PSA_SALARIE','PSA_DATEENTREE'],[Salarie,DateEntree],False);
                        If T <> Nil then
                        begin
                          If T.GetValue('PSA_DATESORTIE') <= IDate1900 then
                          begin
                             If (TobPaie.Detail[p].GetValue('PPU_DATEENTREE') < TobPaie.Detail[p].GetValue('PPU_DATESORTIE')) or (TobPaie.Detail[p].GetValue('PPU_DATESORTIE') <= IDate1900)then
                             begin
                                If p > 0 then
                                begin
                                        if TobPaie.Detail[p-1].GetValue('PPU_LIBELLEEMPLOI') <> null then T.PutValue('PSA_LIBELLEEMPLOI',TobPaie.Detail[p-1].GetValue('PPU_LIBELLEEMPLOI'));
                                        if (TobPaie.Detail[p-1].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p-1].GetValue('PPU_QUALIFICATION') <> null) then
                                        begin
                                             Qualification := TobPaie.Detail[p-1].GetValue('PPU_QUALIFICATION');
                                             Convention := TobPaie.Detail[p-1].GetValue('PPU_CONVENTION');
                                               If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                                               If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                                        end;
                                end;
                                T.PutValue('PSA_DATESORTIE',PlusDate(TobPaie.Detail[p].GetValue('PPU_DATEENTREE'),-1,'J'));
                             end;
                          end;
                          DateEntree := TobPaie.Detail[p].GetValue('PPU_DATEENTREE');
                          DateSortie := TobPaie.Detail[p].GetValue('PPU_DATESORTIE');
                          Etablissement := TobPaie.Detail[p].GetValue('PPU_ETABLISSEMENT');
                          If TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etablissement],False) <> nil then
                          begin
                            ConstruireTob;
                            T.PutValue('TYPEPERSONNE', 'Salarié');
                            RemplirInfosTob(TobSalarie.Detail[i]);
                            if (TobPaie.Detail[p].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p].GetValue('PPU_QUALIFICATION') <> null) then
                            begin
                                 Qualification := TobPaie.Detail[p].GetValue('PPU_QUALIFICATION');
                                 Convention := TobPaie.Detail[p].GetValue('PPU_CONVENTION');
                                   If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                                   If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                            end;
                            T.PutValue('PSA_ETABLISSEMENT', Etablissement);
                            if TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI') <> null then
                            begin
                                 LibelleEmploi := TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI');
                                 T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                            end;
                            T.PutValue('PSA_DATEENTREE', DateEntree);
                            T.PutValue('PSA_DATESORTIE', DateSortie);
                            T.PutValue('PSA_MOTIFSORTIE', '');
                            T.PutValue('MUTATION', '');
                            T.PutValue('PCI_TYPECONTRAT', 'CDI');
                          end;
                        end;
                end
                else
                begin
                    If TobPaie.Detail[p].GetValue('PPU_DATESORTIE') <> DateSortie then
                    begin
                        T := TobEtat.FindFirst(['PSA_SALARIE','PSA_DATEENTREE'],[Salarie,DateEntree],False);
                        If T <> Nil then
                        begin
                          DateSortie := TobPaie.Detail[p].GetValue('PPU_DATESORTIE');
                          T.PutValue('PSA_DATESORTIE',DateSortie);
                        end;
                    end;
                end;
                If p < TobPaie.Detail.Count - 1 then
                begin
                  If TobPaie.detail[p].GetValue('PPU_ETABLISSEMENT') <> TobPaie.Detail[P+1].GetValue('PPU_ETABLISSEMENT') then
                  begin
                    If DateEntree = TobPaie.Detail[p+1].GetValue('PPU_DATEENTREE') then
                    begin
                      T := TobEtat.FindFirst(['PSA_SALARIE','PSA_DATEENTREE'],[Salarie,DateEntree],False);
                      If T <> Nil then
                      begin
                          T.PutValue('PSA_DATESORTIE',TobPaie.Detail[p].GetValue('PPU_DATEFIN'));
                      end;
                      Etablissement := TobPaie.Detail[p+1].GetValue('PPU_ETABLISSEMENT');
                      DateSortie := TobPaie.Detail[p+1].GetValue('PPU_DATESORTIE');
                      DateEntree := TobPaie.Detail[p+1].GetValue('PPU_DATEDEBUT');
                      If TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etablissement],False) <> nil then
                      begin
                          ConstruireTob;
                          T.PutValue('TYPEPERSONNE', 'Salarié');
                          RemplirInfosTob(TobSalarie.Detail[i]);
                          if (TobPaie.Detail[p+1].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p+1].GetValue('PPU_QUALIFICATION') <> null) then
                          begin
                                 Qualification := TobPaie.Detail[p+1].GetValue('PPU_QUALIFICATION');
                                 Convention := TobPaie.Detail[p+1].GetValue('PPU_CONVENTION');
                                   If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                                   If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                          end;
                          T.PutValue('PSA_ETABLISSEMENT', Etablissement);
                          if TobPaie.Detail[p+1].GetValue('PPU_LIBELLEEMPLOI') <> null then
                          begin
                               LibelleEmploi := TobPaie.Detail[p+1].GetValue('PPU_LIBELLEEMPLOI');
                               T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                          end;
                          T.PutValue('PSA_DATEENTREE', DateEntree);
                          T.PutValue('PSA_DATESORTIE', DateSortie);
                          T.PutValue('PSA_MOTIFSORTIE', '');
                          T.PutValue('MUTATION', '');
                          T.PutValue('PCI_TYPECONTRAT', 'CDI');
                      end;
                    end
                    else
                    begin
                      T := TobEtat.FindFirst(['PSA_SALARIE','PSA_DATEENTREE'],[Salarie,DateEntree],False);
                      If T <> Nil then
                      begin
                          If (T.GetValue('PSA_DATESORTIE') <= IDate1900) or (T.GetValue('PSA_DATESORTIE')>TobPaie.Detail[p+1].GetValue('PPU_DATEENTREE')) then
                          T.PutValue('PSA_DATESORTIE',TobPaie.Detail[p].GetValue('PPU_DATEFIN'));
                          if TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI') <> Null then LibelleEmploi := TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI')
                          else LibelleEmploi := '';
                          If LibelleEmploi = '' then LibelleEmploi := TobSalarie.Detail[i].GetValue('PSA_LIBELLEEMPLOI');
                          T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                          if (TobPaie.Detail[p].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p].GetValue('PPU_QUALIFICATION') <> null) then
                          begin
                                 Qualification := TobPaie.Detail[p].GetValue('PPU_QUALIFICATION');
                                 Convention := TobPaie.Detail[p].GetValue('PPU_CONVENTION');
                                   If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                                   If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                          end;
                      end;
                      Etablissement := TobPaie.Detail[p+1].GetValue('PPU_ETABLISSEMENT');
                      DateSortie := TobPaie.Detail[p+1].GetValue('PPU_DATESORTIE');
                      DateEntree := TobPaie.Detail[p+1].GetValue('PPU_DATEENTREE');
                      If TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etablissement],False) <> nil then
                      begin
                          ConstruireTob;
                          T.PutValue('TYPEPERSONNE', 'Salarié');
                          RemplirInfosTob(TobSalarie.Detail[i]);
                          if (TobPaie.Detail[p+1].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p+1].GetValue('PPU_QUALIFICATION') <> null) then
                          begin
                                 Qualification := TobPaie.Detail[p+1].GetValue('PPU_QUALIFICATION');
                                 Convention := TobPaie.Detail[p+1].GetValue('PPU_CONVENTION');
                                   If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                                   If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                          end;
                          T.PutValue('PSA_ETABLISSEMENT', Etablissement);
                          if TobPaie.Detail[p+1].GetValue('PPU_LIBELLEEMPLOI') then
                          begin
                               LibelleEmploi := TobPaie.Detail[p+1].GetValue('PPU_LIBELLEEMPLOI');
                               T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                          end;
                          T.PutValue('PSA_DATEENTREE', DateEntree);
                          T.PutValue('PSA_DATESORTIE', DateSortie);
                          T.PutValue('PSA_MOTIFSORTIE', '');
                          T.PutValue('MUTATION', '');
                          T.PutValue('PCI_TYPECONTRAT', 'CDI');
                      end;
                    end;
                  end;
                end;
                T := TobEtat.FindFirst(['PSA_SALARIE','PSA_DATEENTREE'],[Salarie,DateEntree],False);
                If T <> Nil then
                begin
                        if TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI') <> Null then LibelleEmploi := TobPaie.Detail[p].GetValue('PPU_LIBELLEEMPLOI');
                        If LibelleEmploi = '' then LibelleEmploi := TobSalarie.Detail[i].GetValue('PSA_LIBELLEEMPLOI');
                        T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                        if (TobPaie.Detail[p].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[p].GetValue('PPU_QUALIFICATION') <> null) then
                        begin
                                 Qualification := TobPaie.Detail[p].GetValue('PPU_QUALIFICATION');
                                 Convention := TobPaie.Detail[p].GetValue('PPU_CONVENTION');
                                 If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                                 If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                        end;
                end;
             end;
             T := TobEtat.FindFirst(['PSA_SALARIE','PSA_DATEENTREE'],[Salarie,DateEntree],False);
             If T <> Nil then
             begin
                If TobPaie.Detail[TobPaie.detail.Count - 1].GetValue('PPU_LIBELLEEMPLOI') <> null then
                begin
                     LibelleEmploi := TobPaie.Detail[TobPaie.detail.Count - 1].GetValue('PPU_LIBELLEEMPLOI');
                     T.PutValue('PSA_LIBELLEEMPLOI',LibelleEmploi);
                end;
                if (TobPaie.Detail[TobPaie.detail.Count - 1].GetValue('PPU_CONVENTION') <> null) and (TobPaie.Detail[TobPaie.detail.Count - 1].GetValue('PPU_QUALIFICATION') <> null) then
                 begin
                      Qualification := TobPaie.Detail[TobPaie.detail.Count - 1].GetValue('PPU_QUALIFICATION');
                      Convention := TobPaie.Detail[TobPaie.detail.Count - 1].GetValue('PPU_CONVENTION');
                        If Qualification <> '' then T.PutValue('PSA_QUALIFICATION', Qualification);
                        If Convention <> '' then T.PutValue('PSA_CONVENTION', Convention);
                 end;
             end;
             FreeAndNil(TobPaie);
          end
          else
          begin
                FreeAndNil(TobPaie);
                //Si pas de contrat ni de paie reprise des infos table salarié
                Etablissement := TobSalarie.Detail[c].GetValue('PSA_ETABLISSEMENT');
                If TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etablissement],False) <> nil then
                begin
                  ConstruireTob;
                  LibelleEmploi := TobSalarie.Detail[i].GetValue('PSA_LIBELLEEMPLOI');
                  T.PutValue('TYPEPERSONNE', 'Salarié');
                  RemplirInfosTob(TobSalarie.Detail[i]);
                  T.PutValue('PSA_ETABLISSEMENT', Etablissement);
                  T.PutValue('PSA_LIBELLEEMPLOI', LibelleEmploi);
                  T.PutValue('PSA_DATEENTREE', TobSalarie.Detail[i].GetValue('PSA_DATEENTREE'));
                  T.PutValue('PSA_DATESORTIE', TobSalarie.Detail[i].GetValue('PSA_DATESORTIE'));
                  T.PutValue('PSA_MOTIFSORTIE', TobSalarie.Detail[i].getValue('PSA_MOTIFSORTIE'));
                  T.PutValue('MUTATION', '');
                  T.PutValue('PCI_TYPECONTRAT', 'CDI');
                end;
          end;
        end;
  MoveCurProgressForm(TobSalarie.Detail[i].getValue('PSA_SALARIE')+' '+TobSalarie.Detail[i].getValue('PSA_LIBELLE'));
  end;
  Q := OpenSQL('SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM,PEI_ETABLISSEMENT,' +
    'PSI_ADRESSE1,PSI_ADRESSE2,PSI_ADRESSE3,PSI_VILLE,PSI_CODEPOSTAL,' +
    'PSI_DATENAISSANCE,PSI_SEXE,PSI_NUMEROSS,PSI_NATIONALITE,PEI_LIBELLEEMPLOI,PEI_QUALIFICATION,' +
    'PSI_CARTESEJOUR,PSI_DELIVPAR,PSI_DATEXPIRSEJOUR,' +
    'PEI_AGENCEINTGU,PEI_DEBUTEMPLOI,PEI_FINEMPLOI FROM EMPLOIINTERIM LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PEI_INTERIMAIRE' +
    ' WHERE PEI_ETABLISSEMENT>="'+Etab1+'" AND PEI_ETABLISSEMENT<="'+Etab2+'" AND '+
    'PSI_TYPEINTERIM="INT" AND PEI_DEBUTEMPLOI<="' + UsDatetime(DateLegale) + '" AND PEI_FINEMPLOI>="' + UsDateTime(DateEdition) + '" ORDER BY PSI_INTERIMAIRE', True);
  TobInterim := Tob.create('Les intérimaires', nil, -1);
  TobInterim.LoadDetailDB('INTERIMAIRES', '', '', Q, False);
  Ferme(Q);
  for J := 0 to TobInterim.detail.Count - 1 do
  begin
      ConstruireTob;
      T.PutValue('TYPEPERSONNE', 'Travailleur temporaire');
      T.PutValue('PSA_SALARIE', TobInterim.Detail[j].GetValue('PSI_INTERIMAIRE'));
      T.PutValue('PSA_LIBELLE', TobInterim.Detail[j].GetValue('PSI_LIBELLE'));
      T.PutValue('PSA_PRENOM', TobInterim.Detail[j].GetValue('PSI_PRENOM'));
      T.PutValue('PSA_ETABLISSEMENT', TobInterim.Detail[j].GetValue('PEI_ETABLISSEMENT'));
      T.PutValue('PSA_ADRESSE1', TobInterim.Detail[j].GetValue('PSI_ADRESSE1'));
      T.PutValue('PSA_ADRESSE2', TobInterim.Detail[j].GetValue('PSI_ADRESSE2'));
      T.PutValue('PSA_ADRESSE3', TobInterim.Detail[j].GetValue('PSI_ADRESSE3'));
      T.PutValue('PSA_VILLE', TobInterim.Detail[j].GetValue('PSI_VILLE'));
      T.PutValue('PSA_CODEPOSTAL', TobInterim.Detail[j].GetValue('PSI_CODEPOSTAL'));
      T.PutValue('PSA_DATENAISSANCE', TobInterim.Detail[j].GetValue('PSI_DATENAISSANCE'));
      T.PutValue('PSA_SEXE', TobInterim.Detail[j].GetValue('PSI_SEXE'));
      T.PutValue('PSA_NUMEROSS', TobInterim.Detail[j].GetValue('PSI_NUMEROSS'));
      T.PutValue('PSA_NATIONALITE', TobInterim.Detail[j].GetValue('PSI_NATIONALITE'));
      T.PutValue('PSA_LIBELLEEMPLOI', TobInterim.Detail[j].GetValue('PEI_LIBELLEEMPLOI'));
      T.PutValue('PSA_QUALIFICATION', TobInterim.Detail[j].GetValue('PEI_QUALIFICATION'));
      T.PutValue('PSA_DATEENTREE', TobInterim.Detail[j].GetValue('PEI_DEBUTEMPLOI'));
      T.PutValue('PSA_DATESORTIE', TobInterim.Detail[j].GetValue('PEI_FINEMPLOI'));
      T.PutValue('PSA_CONDEMPLOI', '');
      T.PutValue('PSA_CARTESEJOUR', TobInterim.Detail[j].GetValue('PSI_CARTESEJOUR'));
      T.PutValue('PSA_DELIVPAR', TobInterim.Detail[j].GetValue('PSI_DELIVPAR'));
      T.PutValue('SEJOUR', TobInterim.Detail[j].GetValue('PSI_DATEXPIRSEJOUR'));
      T.PutValue('PCI_TYPECONTRAT', 'CTT');
      T.PutValue('PSA_COEFFICIENT', '');
      T.PutValue('DATELICENCIEMENT', IDate1900);
      T.PutValue('PSA_MOTIFSORTIE', '');
      T.PutValue('MUTATION', '');
      T.PutValue('PSA_CATDADS', '');
      T.PutValue('PSA_DADSCAT', '');
      T.PutValue('PSA_INDICE', '');
      T.PutValue('PSA_CONVENTION','');
      T.PutValue('LIBQUALIF','' );
      TA := TobAgence.FindFirst(['ANN_GUIDPER'], [TobInterim.Detail[j].GetValue('PEI_AGENCEINTGU')], False);
      if TA <> nil then
      begin
        T.PutValue('AGENCE', TA.GetValue('ANN_APNOM'));
        T.PutValue('ADRESSE1', TA.GetValue('ANN_APRUE1'));
        T.PutValue('ADRESSE2', TA.GetValue('ANN_APRUE2') + TA.GetValue('ANN_APRUE2'));
        T.PutValue('CPVILLE', TA.GetValue('ANN_APCPVILLE'));
      end
      else
      begin
        T.PutValue('AGENCE', '');
        T.PutValue('ADRESSE1', '');
        T.PutValue('ADRESSE2', '');
        T.PutValue('CPVILLE', '');
      end;
  end;
  FiniMoveProgressForm;
  FreeAndNil(TobEtab);
  FreeAndNil(TobInterim);
  TobEtat.Detail.Sort('PSA_ETABLISSEMENT;PSA_DATEENTREE');
  For i := 0 to TobEtat.Detail.Count - 1 do
  begin
       Salarie := TobEtat.Detail[i].Getvalue('PSA_SALARIE');
       if TobEtat.Detail[i].Getvalue('PSA_QUALIFICATION') <> Null then Qualification := TobEtat.Detail[i].Getvalue('PSA_QUALIFICATION')
       else qualification := '';
       if TobEtat.Detail[i].Getvalue('PSA_CONVENTION') <> Null then Convention := TobEtat.Detail[i].Getvalue('PSA_CONVENTION')
       else Convention := '';
       LibQualif := RendLibQualif(Qualification,Convention);
       TobEtat.Detail[i].PutValue('LIBQUALIF',LibQualif);
  end;
  FiniMoveProgressForm;

  FreeAndNil(TobAgence);

  {$IFDEF EAGLCLIENT}
    StPages := AglGetCriteres(Pages, FALSE);     //PT2
  {$ENDIF}
  TFQRS1(Ecran).LaTob:= TobEtat;                 //PT2
end;

initialization
  registerclasses([TOF_PGListePersonnelDuAu]);
end.

