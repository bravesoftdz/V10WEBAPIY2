unit UtofPgExportEtemptation;

interface


uses Classes, StdCtrls, SysUtils, Controls,
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  Utof, Utob, Htb97, HEnt1, Hctrls, HmsgBox, Vierge, ParamSoc;
type
  TOF_PGEXPORTETEMPTATION = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
  private
    procedure OnClickRBInitialisation(Sender: TObject);
    procedure OnClickRBWorkFlow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    function PgDoFichierSect(AncDat, DateValidite: TDateTime): Boolean;
    function PgDoFichierHierar(AncDat, DateValidite: TDateTime): Boolean;
    function PgDoFichierMail(AncDat, DateValidite: TDateTime): Boolean;
    function PgDoFichierEmploye(AncDat, DateValidite: TDateTime; ModeWF: Boolean = False): Boolean;
    function PgDoFichierContrat(AncDat, DateValidite: TDateTime; ModeWF: Boolean = False): Boolean;
    function PgDoFichierAbsence(AncDat, DateValidite: TDateTime; ModeWF: Boolean = False): Boolean;

    function OkSuppFichier(FileN, Mess: string): Boolean;

  end;

implementation
uses Entpaie;

type
  SectionHoraire = record { ENREGISTREMENT DE SECTION HORAIRE }
    HorSect: string[3]; { Section Horaire }
    Libelle: string[40]; { Libelle Section Horaire }
    HorCode: string[6]; { Code horaire }
    Rang: string[1]; { Rang }
  end;

type
  Groupe = record { ENREGISTREMENT DE GROUPE }
    Groupe: string[10]; { Groupe }
    Libelle: string[40]; { Libellé Groupe }
  end;

type
  Hierarchie = record { ENREGISTREMENT DE HIERARCHIE }
    Hierar: string[10]; { Hierarchie }
    Libelle: string[40]; { Libellé Hierarchie }
    TypFct: string[1]; { Type de fonction }
  end;

type
  HierarDet = record { ENREGISTREMENT DE HIERARCHIE DETAILLEE }
    Hierar: string[10]; { Hierarchie }
    Matri: string[10]; { Matricule du responsable }
    Final: string[1]; { Fonctionnement }
    Niveau: string[2]; { Niveau }
  end;




type
  Mail = record { ENREGISTREMENT DE ADRESSE MAIL }
    Mods: string[10]; { Mod }
    Matri: string[10]; { Matricule salarié }
    Adresse: string[255]; { Adresse Mail }
  end;

type
  Employe = record { ENREGISTREMENT DE EMPLOYE }
    Matri: string[10]; { Matricule salarié }
    NomPre: string[30]; { Nom prénom }
    DatOuv: string[8]; { Date d'ouverture de l'historique }
    HorSect: string[3]; { Numéro de section horaire }
    SeiTyper: string[1]; { Type de personnel }
    GroupH: string[10]; { Groupe hiérarchique }
    Alias: string[20]; { Alais authentification }
    SeqManager : String[1]; { si employé manager }
  end;

type
  Contrat = record { ENREGISTREMENT DE EMPLOYE }
    Matri: string[10]; { Matricule salarié }
    DatDeb: string[8]; { Date début de contrat }
    DatFin: string[8]; { Date de fin de contrat }
    TypCtra: string[4]; { Type de contrat }
    Libelle: string[30]; { Libellé du contrat }
//    CtrHora: string[6]; { Code horaire contrat }
//    CtrRg: string[2]; { Rang du code horaire }
  end;

type
  Absence = record { ENREGISTREMENT DE ABSENCE }
    Lig: string[3]; {}
    Col: string[3]; {}
    Datejour: string[10]; {}
    Mode: string[8]; {}
    FctPre: string[3]; {}
    FctCur: string[3]; {}
    OptSel: string[1]; {}
    LibOpt: string[20]; {}
    CodSel: string[10]; {}
    LibSel: string[46]; {}
    JourDeb: string[8]; {}
    DateDeb: string[10]; {}
    JourFin: string[8]; {}
    DateFin: string[10]; {}
    Japp: string[7]; {}
    Texte_01: string[20]; {}
    Motif_01: string[4]; {}
    Valoris_01: string[1]; {}
    MotifDur_01: string[7]; {}
    HRaDeb_01: string[6]; {}
    HRaFin_01: string[6]; {}
    Libelle_01: string[8]; {}
    Texte_02: string[20]; {}
    Motif_02: string[4]; {}
    Valoris_02: string[1]; {}
    MotifDur_02: string[7]; {}
    HRaDeb_02: string[6]; {}
    HRaFin_02: string[6]; {}
    Libelle_02: string[8]; {}
    Messages: string[77]; {}
  end;





{ TOF_PGEXPORTETEMPTATION }


procedure TOF_PGEXPORTETEMPTATION.OnArgument(Arguments: string);
var RB: TRadioButton;
  Btn: TToolBarButton97;
begin
  inherited;

  RB := TRadioButton(GetControl('RBINITIALISATION'));
  if Assigned(RB) then RB.OnClick := OnClickRBInitialisation;

  RB := TRadioButton(GetControl('RBWORHFLOW'));
  if Assigned(RB) then RB.OnClick := OnClickRBWorkFlow;

  Btn := TToolBarButton97(GetControl('BValider'));
  if Assigned(Btn) then Btn.OnClick := BValiderClick;

//  if VH_PAie.PGETempDateVal > idate1900 then
  SetControlText('ANCDATEVAL', DateToStr(VH_PAie.PGETempDateVal));

  if Assigned(THCheckBox(GetControl('RBTRAVAILN2'))) then 
  THCheckBox(GetControl('RBTRAVAILN2')).Caption := 'Sections : '+VH_Paie.PGLibelleOrgStat2;


end;



procedure TOF_PGEXPORTETEMPTATION.OnClickRBInitialisation(Sender: TObject);
begin
//SetControlProperty('GBINIT','Enabled',TRadioButton(Sender).Checked = True);
//SetControlProperty('GBWORKFLOW','Enabled',Not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('DATEVALIDITE', TRadioButton(Sender).Checked = True);
  SetControlEnabled('RBSALARIE', not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('RBCONTRAT', not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('RBABSENCE', not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('RBMAIL', not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('RBHIERACH', not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('RBTRAVAILN2', not (TRadioButton(Sender).Checked = True));

end;

procedure TOF_PGEXPORTETEMPTATION.OnClickRBWorkFlow(Sender: TObject);
begin
  SetControlEnabled('DATEVALIDITE', not (TRadioButton(Sender).Checked = True));
  SetControlEnabled('RBSALARIE', TRadioButton(Sender).Checked = True);
  SetControlEnabled('RBCONTRAT', TRadioButton(Sender).Checked = True);
  SetControlEnabled('RBABSENCE', TRadioButton(Sender).Checked = True);
  SetControlEnabled('RBMAIL', TRadioButton(Sender).Checked = True);
  SetControlEnabled('RBHIERACH', TRadioButton(Sender).Checked = True);
  SetControlEnabled('RBTRAVAILN2', TRadioButton(Sender).Checked = True);
end;


procedure TOF_PGEXPORTETEMPTATION.BValiderClick(Sender: TObject);
var
  RB: TRadioButton;
  DateValidite, AncDatVal: TDateTime;
  CHBX : THCheckBox;
  OkOk : Boolean;
begin
  if IsValidDate(GetControlText('DATEVALIDITE')) then
    DateValidite := StrToDate(GetControlText('DATEVALIDITE'))
  else
    exit;

  if IsValidDate(GetControlText('ANCDATEVAL')) then
    AncDatVal := StrToDate(GetControlText('ANCDATEVAL'))
  else
    exit;

  OkOk := False;
  { Phase d'initialisation }
  RB := TRadioButton(GetControl('RBINITIALISATION'));
  if Assigned(RB) then
    if RB.Checked = True then
    begin
      OkOk := True;
      if not PgDoFichierSect(AncDatVal, DateValidite) then exit;
      if not PgDoFichierHierar(AncDatVal, DateValidite) then exit;
      if not PgDoFichierMail(AncDatVal, DateValidite) then exit;
      if not PgDoFichierEmploye(AncDatVal, DateValidite) then exit;
      if not PgDoFichierContrat(AncDatVal, DateValidite) then exit;
      if not PgDoFichierAbsence(AncDatVal, DateValidite) then exit;
    end;

  { Phase WorkFlow }
  RB := TRadioButton(GetControl('RBWORHFLOW'));
  if Assigned(RB) then
    if RB.Checked = True then
    begin
      OkOk := False;
      CHBX := THCheckBox(GetControl('RBSALARIE'));
      if (Assigned(CHBX)) and (CHBX.Checked) then
        if PgDoFichierEmploye(AncDatVal, DateValidite, True) then OkOk := True;

      CHBX := THCheckBox(GetControl('RBCONTRAT'));
      if (Assigned(CHBX)) and (CHBX.Checked) then
        if PgDoFichierContrat(AncDatVal, DateValidite, True) then OkOk := True;

      CHBX := THCheckBox(GetControl('RBABSENCE'));
      if (Assigned(CHBX)) and (CHBX.Checked) then
        if PgDoFichierAbsence(AncDatVal, DateValidite, True)  then OkOk := True;

      CHBX := THCheckBox(GetControl('RBMAIL'));
      if (Assigned(CHBX)) and (CHBX.Checked) then
        if PgDoFichierMail(AncDatVal, DateValidite) then OkOk := True;

      CHBX := THCheckBox(GetControl('RBHIERACH'));
      if (Assigned(CHBX)) and (CHBX.Checked) then
        if PgDoFichierHierar(AncDatVal, DateValidite) then OkOk := True;

      CHBX := THCheckBox(GetControl('RBTRAVAILN2'));
      if (Assigned(CHBX)) and (CHBX.Checked) then
        if PgDoFichierSect(AncDatVal, DateValidite) then OkOk := True;    
    end;

    if OkOk then
      Begin
      PgiInfo('Traitement terminé. Génération des fichiers sous ' + VH_PAIE.PGRepPublic + '.', TFVierge(Ecran).caption);

      if PgiAsk('Voulez-vous mettre à jour la date d''arrêté historisée?', TFVierge(Ecran).caption) = MrYes then
        begin
        SetParamSoc('SO_PGETEMPDATEVAL', DateValidite);
        VH_PAIE.PGETempDateVal := DateValidite;
        SetControlText('ANCDATEVAL', DateToStr(DateValidite));
        end;
      end;



end;

function TOF_PGEXPORTETEMPTATION.PgDoFichierSect(AncDat, DateValidite: TDateTime): Boolean;
var
  Q: TQuery;
  St, FileN : string;
  F : TextFile;
  Tob_Fichier: Tob;
  SH: SectionHoraire;
  i: integer;
begin
  Result := True;
  Tob_Fichier := nil;
  St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="PST"';
  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Fichier := Tob.create('le fichier', nil, 1);
    Tob_Fichier.LoadDetailDB('le fichier', '', '', Q, False);
  end
  else
  begin
    Ferme(Q);
    PgiBox('Aucune donnée section horaire à exporter.', TFVierge(Ecran).caption);
    Exit;
  end;
  Ferme(Q);

  if Assigned(Tob_Fichier) then
  begin
      { Recherche du fichier des sections horaires }
    FileN := VH_Paie.PGRepPublic + '\sech.csv';
    if not OkSuppFichier(FileN, 'Voulez-vous supprimer le fichier des sections horaires?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;


    St := 'HORSECT;LIBELLE;HORCODE;RANG';
    writeln(F, St);

    for i := 0 to Tob_Fichier.detail.count - 1 do
    begin
        { ENREGISTREMENT DE SECTION HORAIRE }
      SH.HorSect := Copy(Tob_Fichier.detail[i].GetValue('CC_CODE'), 1, SizeOf(SH.HorSect)); { Section Horaire }
      SH.Libelle := Copy(Tob_Fichier.detail[i].GetValue('CC_LIBELLE'), 1, SizeOf(SH.Libelle)); { Libelle Section Horaire }
      SH.HorCode := 'STP'; { Code horaire }
      SH.Rang := '1'; { Rang }
      St := SH.HorSect + ';' + SH.Libelle + ';' + SH.HorCode + ';' + SH.Rang;
      writeln(F, St);
    end;
    CloseFile(F);
  end;
  if Assigned(Tob_Fichier) then FreeAndNil(Tob_Fichier);
end;



function TOF_PGEXPORTETEMPTATION.PgDoFichierHierar(AncDat, DateValidite: TDateTime): Boolean;
var
  Q: TQuery;
  St, FileN, FileN2, FileN3: string;
  F, F2, F3: TextFile;
  Tob_Fichier: Tob;
  HR: Hierarchie;
  HD: HierarDet;
  GR: Groupe;
  i: integer;
begin
  Result := True;
  Tob_Fichier := nil;
  St := 'SELECT PGS_CODESERVICE,PGS_NOMSERVICE||" - "||PSA_LIBELLE||" "||'+
        'SUBSTRING(PSA_PRENOM,1,1)||""||"." AS NOMSERVICE,PGS_RESPONSABS '+
        'FROM SERVICES,SALARIES WHERE PGS_RESPONSABS=PSA_SALARIE';
  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Fichier := Tob.create('le fichier', nil, 1);
    Tob_Fichier.LoadDetailDB('le fichier', '', '', Q, False);
  end
  else
  begin
    Ferme(Q);
    PgiBox('Aucune donnée hiérarchie à exporter.', TFVierge(Ecran).caption);
    Exit;
  end;
  Ferme(Q);

  if Assigned(Tob_Fichier) then
  begin
      { Recherche du fichier des hiérarchies }
    FileN := VH_Paie.PGRepPublic + '\sfphierg.csv';
    if not OkSuppFichier(FileN, 'Voulez-vous supprimer le fichier des hiérarchies?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

      { Recherche du fichier des hiérarchies détaillées }
    FileN2 := VH_Paie.PGRepPublic + '\sfphierd.csv';
    if not OkSuppFichier(FileN2, 'Voulez-vous supprimer le fichier des responsables hiérarchiques?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;


      { Recherche du fichier des groupes }
    FileN3 := VH_Paie.PGRepPublic + '\sfpgrps.csv';
    if not OkSuppFichier(FileN3, 'Voulez-vous supprimer le fichier des groupes?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;


    AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    AssignFile(F2, FileN2);
{$I-}ReWrite(F2);
{$I+}
    if IoResult <> 0 then
    begin
      CloseFile(F);
      FreeAndNil(Tob_Fichier);
      PGIBox('Fichier inaccessible : ' + FileN2, 'Abandon du traitement');
      Result := False;
      Exit;
    end;

    AssignFile(F3, FileN3);
{$I-}ReWrite(F3);
{$I+}
    if IoResult <> 0 then
    begin
      CloseFile(F);
      CloseFile(F2);
      PGIBox('Fichier inaccessible : ' + FileN3, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;


    St := 'HIERAR;LIBELLE;TYPFCT';
    writeln(F, St);
    St := 'HIERAR;MATRI;FINAL;NIVEAU';
    writeln(F2, St);
    St := 'GROUPE;LIBELLE';
    writeln(F3, St);        


    for i := 0 to Tob_Fichier.detail.count - 1 do
    begin
        { ENREGISTREMENT DE HIERARCHIE }
      HR.Hierar := Copy(Tob_Fichier.detail[i].GetValue('PGS_CODESERVICE'), 1, SizeOf(HR.Hierar)); { Hierarchie }
      HR.Libelle := Copy(Tob_Fichier.detail[i].GetValue('NOMSERVICE'), 1, SizeOf(HR.Libelle)); { Libelle Hierarchie }
      HR.TypFct := '3'; { Type fonction }
      St := HR.Hierar + ';' + HR.Libelle + ';' + HR.TypFct;
      writeln(F, St);
        { ENREGISTREMENT DE HIERARCHIE DETAILLE}
      HD.Hierar := Copy(Tob_Fichier.detail[i].GetValue('PGS_CODESERVICE'), 1, SizeOf(HD.Hierar)); { Hierarchie }
      HD.Matri := Copy(Tob_Fichier.detail[i].GetValue('PGS_RESPONSABS'), 1, SizeOf(HD.Matri)); { Responsable }
      HD.Final := '1'; { Fonctionnement }
      HD.Niveau := '0'; { Niveau }
      St := HD.Hierar + ';' + HD.Matri + ';' + HD.Final + ';' + HD.Niveau;
      writeln(F2, St);
              { ENREGISTREMENT DE GROUPE }
      GR.Groupe := Copy(Tob_Fichier.detail[i].GetValue('PGS_CODESERVICE'), 1, SizeOf(GR.Groupe)); { Groupe }
      GR.Libelle := Copy(Tob_Fichier.detail[i].GetValue('NOMSERVICE'), 1, SizeOf(GR.Libelle)); { Libelle }
      St := GR.Groupe + ';' + GR.Libelle;
      writeln(F3, St);
    end;
    CloseFile(F);
    CloseFile(F2);
    CloseFile(F3);
  end;
  if Assigned(Tob_Fichier) then FreeAndNil(Tob_Fichier);
end;


function TOF_PGEXPORTETEMPTATION.PgDoFichierMail(AncDat, DateValidite: TDateTime): Boolean;
var
  Q: TQuery;
  St, FileN: string;
  F: TextFile;
  Tob_Fichier: Tob;
  MS: Mail;
  i: integer;
begin
  Result := True;
  Tob_Fichier := nil;
  if ( NOT (GetParamSocSecur ('SO_IFDEFCEGID', FALSE))) then St := 'SELECT DISTINCT US_AUXILIAIRE,US_EMAIL FROM UTILISAT WHERE US_AUXILIAIRE<>"" '
  else St := 'SELECT DISTINCT US_AUXILIAIRE,US_EMAIL FROM PGSICUTILISAT WHERE US_AUXILIAIRE<>"" ';
   //PSE_SALARIE,PSE_EMAILPROF FROM DEPORTSAL';
  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Fichier := Tob.create('le fichier', nil, 1);
    Tob_Fichier.LoadDetailDB('le fichier', '', '', Q, False);
  end
  else
  begin
    Ferme(Q);
    PgiBox('Aucune donnée hiérarchie détaillée à exporter.', TFVierge(Ecran).caption);
    Exit;
  end;
  Ferme(Q);

  if Assigned(Tob_Fichier) then
  begin
      { Recherche du fichier des mails }
    FileN := VH_Paie.PGRepPublic + '\mail.csv';
    if not OkSuppFichier(FileN, 'Voulez-vous supprimer le fichier des mails?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    St := 'MOD;MATRI;ADRESSE';
    writeln(F, St);

    for i := 0 to Tob_Fichier.detail.count - 1 do
    begin
      { ENREGISTREMENT DE MAIL }
      MS.Mods := '';
      MS.Matri := Copy(Tob_Fichier.detail[i].GetValue('US_AUXILIAIRE'), 1, SizeOf(MS.Matri)); { Hierarchie }
      MS.Adresse := Copy(Tob_Fichier.detail[i].GetValue('US_EMAIL'), 1, SizeOf(MS.Adresse)); { Libelle Hierarchie }
      St := MS.Mods + ';' + MS.Matri + ';' + MS.Adresse;
      writeln(F, St);
    end;
    CloseFile(F);
  end;
  if Assigned(Tob_Fichier) then FreeAndNil(Tob_Fichier);
end;


function TOF_PGEXPORTETEMPTATION.PgDoFichierEmploye(AncDat, DateValidite: TDateTime; ModeWF: Boolean = False): Boolean;
var
  Q: TQuery;
  St, FileN: string;
  F: TextFile;
  Tob_Fichier, Tob_Resp: Tob;
  EMP: Employe;
  i: integer;
  LaTable : String;
begin
  Result := True;
  Tob_Fichier := nil;
  St := '';
  if ModeWF then
    St := 'AND PSA_DATEMODIF >="' + USDateTime(EncodeDate(2006,11,28)) + '" ';
  if ( NOT (GetParamSocSecur ('SO_IFDEFCEGID', FALSE))) then LaTable := 'UTILISAT'
  else LaTable := 'PGSICUTILISAT';
  St := 'SELECT DISTINCT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_TRAVAILN2,PSA_BOOLLIBRE1,PSE_CODESERVICE,US_ABREGE ' +
    'FROM SALARIES,DEPORTSAL,'+LaTable+' WHERE PSA_SALARIE=PSE_SALARIE AND US_AUXILIAIRE=PSA_SALARIE ' +
    'AND ((PSA_SUSPENSIONPAIE = "X" AND PSA_SORTIEDEFINIT <> "X") '+
    'OR (PSA_DATEENTREE<="' + USDateTime(DateValidite) + '" ' +
    'AND ( PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="' + USDateTime(Idate1900) + '" ' +
    'OR PSA_DATESORTIE>="' + USDateTime(DateValidite) + '") ' + St+ '))';
  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Fichier := Tob.create('le fichier', nil, 1);
    Tob_Fichier.LoadDetailDB('le fichier', '', '', Q, False);
  end
  else
  begin
    Ferme(Q);
    PgiBox('Aucune donnée employé à exporter.', TFVierge(Ecran).caption);
    Exit;
  end;
  Ferme(Q);

  if Assigned(Tob_Fichier) then
  begin
      { Recherche du fichier des employés }
    FileN := VH_Paie.PGRepPublic + '\empl.csv';
    if not OkSuppFichier(FileN, 'Voulez-vous supprimer le fichier des employés?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    St := 'SELECT PGS_RESPONSABS FROM SERVICES WHERE PGS_RESPONSABS<>""';
    Q := OpenSql(St, True);
    if not Q.Eof then
      begin
      Tob_Resp := Tob.create('les resp', nil, 1);
      Tob_Resp.LoadDetailDB('les resp', '', '', Q, False);
      end
    else
       Tob_Resp := Nil;
    Ferme(Q);

    St := 'MATRI;NOMPRE;DATOUV;HORSECT;SEITYPER;GROUPH;ALIAS;SEQMANAGER';
    writeln(F, St);

    for i := 0 to Tob_Fichier.detail.count - 1 do
    begin
        { ENREGISTREMENT DE EMPLOYE }
      EMP.Matri := Copy(Tob_Fichier.detail[i].GetValue('PSA_SALARIE'), 1, SizeOf(EMP.Matri));
      St := Tob_Fichier.detail[i].GetValue('PSA_LIBELLE') + ' ' + Tob_Fichier.detail[i].GetValue('PSA_PRENOM');
      EMP.NomPre := Copy(St, 1, SizeOf(EMP.NomPre));
      EMP.DatOuv := FormatDateTime('dd/mm/yy', Tob_Fichier.detail[i].GetValue('PSA_DATEENTREE'));
      EMP.HorSect := Copy(Tob_Fichier.detail[i].GetValue('PSA_TRAVAILN2'), 1, SizeOf(EMP.HorSect));
      if Tob_Fichier.detail[i].GetValue('PSA_BOOLLIBRE1') = 'X' then EMP.SeiTyper := '1'
      else EMP.SeiTyper := '2';
      EMP.GroupH := Copy(Tob_Fichier.detail[i].GetValue('PSE_CODESERVICE'), 1, SizeOf(EMP.GroupH));
      EMP.Alias := Copy(Tob_Fichier.detail[i].GetValue('US_ABREGE'), 1, SizeOf(EMP.Alias));
      St := '0';
      if Assigned(Tob_Resp) then
        if Assigned(Tob_Resp.FindFirst(['PGS_RESPONSABS'],[EMP.Matri],False)) then
          St := '1';
      EMP.SeqManager := St;
      St := EMP.Matri + ';' + EMP.NomPre + ';' + EMP.DatOuv + ';' + EMP.HorSect +
            ';' + EMP.SeiTyper + ';' + EMP.GroupH + ';' + EMP.Alias + ';' + EMP.SeqManager;
      writeln(F, St);
    end;
    CloseFile(F);
  end;
  if Assigned(Tob_Fichier) then FreeAndNil(Tob_Fichier);
end;


function TOF_PGEXPORTETEMPTATION.PgDoFichierContrat(AncDat, DateValidite: TDateTime; ModeWF: Boolean = False): Boolean;
var
  Q: TQuery;
  St, FileN, Salarie: string;
  F: TextFile;
  Tob_Fichier, Tob_Sal, T1: Tob;
  CT: Contrat;
  i: integer;
begin
  Result := True;
  Tob_Fichier := nil;
  Tob_Sal := nil;
  St := '';
  if ModeWF then
    St := 'AND(PCI_DEBUTCONTRAT>="' + USDateTime(AncDat) + '" ' +
      'OR (PCI_FINCONTRAT>="' + USDateTime(AncDat) + '" ' +
      'AND PCI_FINCONTRAT<="' + USDateTime(DateValidite) + '")) '
  else
    St := 'AND PCI_DEBUTCONTRAT<="' + USDateTime(DateValidite) + '" ' +
      'AND ( PCI_FINCONTRAT IS NULL OR PCI_FINCONTRAT<="' + USDateTime(Idate1900) + '" ' +
      'OR PCI_FINCONTRAT>="' + USDateTime(DateValidite) + '") ';

  St := 'SELECT PCI_SALARIE,PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,CO_LIBELLE ' +
    'FROM CONTRATTRAVAIL,COMMUN ' +
    'WHERE PCI_TYPECONTRAT=CO_CODE AND CO_TYPE="PCT" ' + St +
    'ORDER BY PCI_SALARIE,PCI_DEBUTCONTRAT DESC';
  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Fichier := Tob.create('le fichier', nil, 1);
    Tob_Fichier.LoadDetailDB('le fichier', '', '', Q, False);
  end
  else
  begin
    Ferme(Q);
    PgiBox('Aucune donnée contrat à exporter.', TFVierge(Ecran).caption);
    Exit;
  end;
  Ferme(Q);

  St := '';
  if ModeWF then
    St := 'AND PSA_DATEENTREE >="' + USDateTime(AncDat) + '" ';

  St := 'SELECT DISTINCT PSA_SALARIE,PSA_DATEENTREE,PSA_DATESORTIE ' +
    'FROM SALARIES WHERE PSA_DATEENTREE<="' + USDateTime(DateValidite) + '" ' +
    'AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="' + USDateTime(Idate1900) + '" ' +
    'OR PSA_DATESORTIE>="' + USDateTime(DateValidite) + '") ' + St;
  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Sal := Tob.create('les sal', nil, 1);
    Tob_Sal.LoadDetailDB('les sal', '', '', Q, False);
  end;
  Ferme(Q);


  if Assigned(Tob_Fichier) then
  begin
      { Recherche du fichier des contrats }
    FileN := VH_Paie.PGRepPublic + '\ctra.csv';
    if not OkSuppFichier(FileN, 'Voulez-vous supprimer le fichier des contrats?') then
    begin
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    if not Assigned(Tob_Sal) then
    begin
      PGIBox('Aucun salarié saisi.', 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;

    St := 'MATRI;DATDEB;DATFIN;TYPCTRA;LIBELLE'; // CTRHORA;CTRRG';
    writeln(F, St);

    Salarie := '';
    for i := 0 to Tob_Sal.detail.count - 1 do
    begin
      CT.Matri := Copy(Tob_Sal.detail[i].GetValue('PSA_SALARIE'), 1, SizeOf(CT.Matri));
//      CT.CtrHora := 'SED';
//      CT.CtrRg := '01';
      Salarie := Tob_Sal.detail[i].GetValue('PSA_SALARIE');
      T1 := Tob_Fichier.findFirst(['PCI_SALARIE'], [Salarie], False);
      if Assigned(T1) then
      begin
           { ENREGISTREMENT DE CONTRAT }
        CT.DatDeb := FormatDateTime('dd/mm/yy', T1.GetValue('PCI_DEBUTCONTRAT'));
        if T1.GetValue('PCI_FINCONTRAT') <= Idate1900 then CT.DatFin := ''
        else CT.DatFin := FormatDateTime('dd/mm/yy', T1.GetValue('PCI_FINCONTRAT'));
        CT.TypCtra := Copy(T1.GetValue('PCI_TYPECONTRAT'), 1, SizeOf(CT.TypCtra));
        CT.Libelle := Copy(T1.GetValue('CO_LIBELLE'), 1, SizeOf(CT.Libelle));
      end
      else
      begin
        CT.DatDeb := FormatDateTime('dd/mm/yy', Tob_Sal.detail[i].GetValue('PSA_DATEENTREE'));
        if Tob_Sal.detail[i].GetValue('PSA_DATESORTIE') <= Idate1900 then CT.DatFin := ''
        else CT.DatFin := FormatDateTime('dd/mm/yy', Tob_Sal.detail[i].GetValue('PSA_DATESORTIE'));
        CT.TypCtra := 'CDI';
        CT.Libelle := Copy('Contrat durée indéterminée', 1, SizeOf(CT.Libelle));
      end;

      St := CT.Matri + ';' + CT.DatDeb + ';' + CT.DatFin + ';' + CT.TypCtra + ';' + CT.Libelle; //  + ';' + CT.CtrHora + ';' + CT.CtrRg;
      writeln(F, St);
    end;
    CloseFile(F);
  end;
  if Assigned(Tob_Fichier) then FreeAndNil(Tob_Fichier);
  if Assigned(Tob_Sal) then FreeAndNil(Tob_Sal);

end;


function TOF_PGEXPORTETEMPTATION.PgDoFichierAbsence(AncDat, DateValidite: TDateTime; ModeWF: Boolean = False): Boolean;
var
  Q: TQuery;
  St, FileN: string;
  F: TextFile;
  Tob_Fichier, TFille: Tob;
  ABS: Absence;
  i: integer;
begin
  Result := True;
  Tob_Fichier := nil;
  St := '';
  if ModeWF then
    St := 'AND PCN_DATEMODIF >= "' + USDateTime(AncDat) + '" ' +
      'AND  PCN_DATEMODIF<= "' + USDateTime(DateValidite + 1 ) + '"  '+
      'AND  PCN_DATEFINABS<= "' + USDateTime(FindeMois(DateValidite)) + '" '
  else
    St := 'AND (PCN_DATEDEBUTABS >= "' + USDateTime(DebutdeMois(DateValidite)) + '" ' +
      'OR PCN_DATEFINABS>= "' + USDateTime(DebutdeMois(DateValidite)) + '") ';

 if VH_Paie.PGEcabBaseDeporte then
    St := St + ' AND PCN_EXPORTOK<>"X" AND PCN_VALIDRESP="VAL" ';


  St := 'SELECT * FROM ABSENCESALARIE ' +
    'WHERE (PCN_TYPEMVT="ABS" ' +
    'OR (PCN_PERIODECP < 2 AND PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE="-")) ' +
    'AND PCN_DATEFINABS >= "' + USDateTime(EncodeDate(2007,1,1)) + '" '+
    'AND PCN_VALIDRESP<>"ATT" ' + St;

  Q := OpenSql(St, True);
  if not Q.Eof then
  begin
    Tob_Fichier := Tob.create('le fichier', nil, 1);
    Tob_Fichier.LoadDetailDB('le fichier', '', '', Q, False);
  end
  else
  begin
    Ferme(Q);
    PgiBox('Aucune donnée absence à exporter.', TFVierge(Ecran).caption);
    Exit;
  end;
  Ferme(Q);

  if Assigned(Tob_Fichier) then
  begin
      { Recherche du fichier des absences }
    FileN := VH_Paie.PGRepPublic + '\HOINACO.dat';
    if not VH_Paie.PGEcabBaseDeporte then
       if not OkSuppFichier(FileN, 'Voulez-vous supprimer le fichier des absences?') then
          begin
          FreeAndNil(Tob_Fichier);
          Result := False;
          Exit;
          end;

    AssignFile(F, FileN);
    if (VH_Paie.PGEcabBaseDeporte) AND FileExists(FileN) then
    {$I-}Append(F) {$I+}
    else
    {$I-}ReWrite(F);{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      FreeAndNil(Tob_Fichier);
      Result := False;
      Exit;
    end;
      { Définition des mvts en création }
    for i := 0 to Tob_Fichier.detail.count - 1 do
    begin
      TFille := Tob_Fichier.detail[i];
      if TFille.GetValue('PCN_ETATPOSTPAIE') = 'NAN' then { Mvt annulé dans la paie }
        TFille.AddChampSupValeur('MODE', 'SUP')
      else
        if TFille.GetValue('PCN_VALIDRESP') <> 'VAL' then { Mvt annulé dans econges }
          TFille.AddChampSupValeur('MODE', 'SUP')
        else
          TFille.AddChampSupValeur('MODE', 'CRE');
      TFille.AddChampSupValeur('CONGES', UpperCase(TFille.GetValue('PCN_TYPECONGE')));
      if UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'PRI' then TFille.PutValue('CONGES', 'CPP')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'RTS') then TFille.PutValue('CONGES', 'RTT')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'RTT') then TFille.PutValue('CONGES', 'RTT')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'AT') then TFille.PutValue('CONGES', 'MAL')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'MAL') then TFille.PutValue('CONGES', 'MAL')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'MAT') then TFille.PutValue('CONGES', 'MAL')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ATJ') then TFille.PutValue('CONGES', 'MAL')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'PAT') then TFille.PutValue('CONGES', 'MAL')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ABA') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ABD') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ABM') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ABP') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'AEM') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ALL') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ABP') then TFille.PutValue('CONGES', 'EVF')
      else if (UpperCase(TFille.GetValue('PCN_TYPECONGE')) = 'ASS') then TFille.PutValue('CONGES', 'ASS')
      else TFille.PutValue('CONGES', 'EVF')
    end;

    for i := 0 to Tob_Fichier.detail.count - 1 do
    begin
        { ENREGISTREMENT DE ABSENCE }
      ABS.Lig := Format_String(Tob_Fichier.detail[i].GetValue('MODE'), SizeOf(ABS.Lig));
      ABS.Col := StringOfChar(' ', SizeOf(ABS.Col));
      ABS.Datejour := StringOfChar(' ', SizeOf(ABS.Datejour));
      ABS.Mode := StringOfChar(' ', SizeOf(ABS.Mode));
      ABS.FctPre := StringOfChar(' ', SizeOf(ABS.FctPre));
      ABS.FctCur := StringOfChar(' ', SizeOf(ABS.FctCur));
      ABS.OptSel := StringOfChar(' ', SizeOf(ABS.OptSel));
      ABS.LibOpt := StringOfChar(' ', SizeOf(ABS.LibOpt));
      ABS.CodSel := Format_String(Tob_Fichier.detail[i].GetValue('PCN_SALARIE'), SizeOf(ABS.CodSel));
      ABS.LibSel := StringOfChar(' ', SizeOf(ABS.LibSel));
      ABS.JourDeb := StringOfChar(' ', SizeOf(ABS.JourDeb));
      ABS.DateDeb := FormatDateTime('dd/mm/yyyy', Tob_Fichier.detail[i].GetValue('PCN_DATEDEBUTABS'));
      ABS.JourFin := StringOfChar(' ', SizeOf(ABS.JourFin));
      ABS.DateFin := FormatDateTime('dd/mm/yyyy', Tob_Fichier.detail[i].GetValue('PCN_DATEFINABS'));
      ABS.Japp := StringOfChar(' ', SizeOf(ABS.Japp));
      ABS.Texte_01 := Format_String(Tob_Fichier.detail[i].GetValue('CONGES'), SizeOf(ABS.Texte_01));
      ABS.Motif_01 := StringOfChar(' ', SizeOf(ABS.Motif_01));
      if (Tob_Fichier.detail[i].GetValue('PCN_DEBUTDJ') = Tob_Fichier.detail[i].GetValue('PCN_FINDJ'))
        and (Tob_Fichier.detail[i].GetValue('PCN_DEBUTDJ') = 'MAT') then
        ABS.Valoris_01 := 'M'
      else
        if (Tob_Fichier.detail[i].GetValue('PCN_DEBUTDJ') = Tob_Fichier.detail[i].GetValue('PCN_FINDJ'))
          and (Tob_Fichier.detail[i].GetValue('PCN_DEBUTDJ') = 'PAM') then
          ABS.Valoris_01 := 'A'
        else
          ABS.Valoris_01 := 'J';
      ABS.MotifDur_01 := StringOfChar(' ', SizeOf(ABS.MotifDur_01));
      ABS.HRaDeb_01 := StringOfChar(' ', SizeOf(ABS.HRaDeb_01));
      ABS.HRaFin_01 := StringOfChar(' ', SizeOf(ABS.HRaFin_01));
      ABS.Libelle_01 := StringOfChar(' ', SizeOf(ABS.Libelle_01));
      ABS.Texte_02 := StringOfChar(' ', SizeOf(ABS.Texte_02));
      ABS.Motif_02 := StringOfChar(' ', SizeOf(ABS.Motif_02));
      ABS.Valoris_02 := StringOfChar(' ', SizeOf(ABS.Valoris_02));
      ABS.MotifDur_02 := StringOfChar(' ', SizeOf(ABS.MotifDur_02));
      ABS.HRaDeb_02 := StringOfChar(' ', SizeOf(ABS.HRaDeb_02));
      ABS.HRaFin_02 := StringOfChar(' ', SizeOf(ABS.HRaFin_02));
      ABS.Libelle_02 := StringOfChar(' ', SizeOf(ABS.Libelle_02));
      ABS.Messages := StringOfChar(' ', SizeOf(ABS.Messages));
      St := ABS.Lig + ABS.Col + ABS.Datejour + ABS.Mode + ABS.FctPre + ABS.FctCur +
        ABS.OptSel + ABS.LibOpt + ABS.CodSel + ABS.LibSel +
        ABS.JourDeb + ABS.DateDeb + ABS.JourFin + ABS.DateFin + ABS.Japp +
        ABS.Texte_01 + ABS.Motif_01 + ABS.Valoris_01 + ABS.MotifDur_01 +
        ABS.HRaDeb_01 + ABS.HRaFin_01 + ABS.Libelle_01 + ABS.Texte_02 +
        ABS.Motif_02 + ABS.Valoris_02 + ABS.MotifDur_02 + ABS.HRaDeb_02 +
        ABS.HRaFin_02 + ABS.Libelle_02 + ABS.Messages;
      write(F, St);
    end;
    if VH_Paie.PGEcabBaseDeporte then
    {$I-}Flush(F);{$I+}
    CloseFile(F);
  end;
  if Assigned(Tob_Fichier) then FreeAndNil(Tob_Fichier);
end;


function TOF_PGEXPORTETEMPTATION.OkSuppFichier(FileN, Mess: string): Boolean;
begin
  Result := True;
  if FileExists(FileN) then
  begin
    if PgiAsk(TraduireMemoire(Mess), TFVierge(Ecran).caption) = MrYes then
    begin
      DeleteFile(PChar(FileN));
      Result := True;
    end
    else
      Result := False;
  end;

end;



initialization
  registerclasses([TOF_PGEXPORTETEMPTATION]);
end.

