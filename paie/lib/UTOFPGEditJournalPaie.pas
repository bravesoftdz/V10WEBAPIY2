{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : Edition du journal de paie
Mots clefs ... : PAIE;CP
*****************************************************************
PT1 18/12/2001 SB V571 Edition sous format liste d'exportation
PT2 29/01/2002 SB V571 Sélection alphanumérique du code salarié
PT3 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT4 20/11/2002 SB V591 Traitements Etats chainés
PT5 24/12/2002 SB V591 FQ 10244 Ajout rupture période
PT6 01/04/2003 SB V_42 Optimisation des temps de traitement
PT7 02/10/2003 SB V_42 Affichage des ongles si gestion paramsoc des combos libres
PT8 10/12/2003 SB V_50 FQ 11005 Prise en compte de la date de fin pour la session de paie
PT9 11/12/2003 SB V_50 FQ 10982 Suppression option monnaie inversée
PT10 18/06/2004 PH V_50 FQ 11164 Base SS pratiquée au lieu de base SS
PT11 18/10/2005 SB V_65 FQ 12007 Ajout critère bullcompl
}
unit UTOFPGEditJournalPaie;

interface

uses StdCtrls,  Classes,  sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc, HQry, UTOFPGEtats;


type
  TOF_PGJOURPAIE_ETAT = class(TOF_PGEtats)
  private
    AjoutChamp, AjoutTri, Origine: string;
    DateDebut, DateFin: TDateTime; //PT4
    procedure DateElipsisclick(Sender: TObject);
    //       procedure MonnaieInverse(Sender: TObject); PT9
    procedure ChangeLieuTravail(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure ExitEdit(Sender: TObject);
    procedure ExportFiche(Sender: TObject); //PT1

  end;
implementation

uses PgEditOutils,PgEditOutils2, EntPaie, PGoutils2, PGEdtEtat;


procedure TOF_PGJOURPAIE_ETAT.OnArgument(Arguments: string);
var
  Check: TCheckBox;
  CDd, CDf, Defaut: THEdit;
  Min, Max, DebPer, FinPer, ExerPerEncours: string;
begin
  inherited;
  //DEB PT4
  DateDebut := idate1900;
  DateFin := idate1900;
  if (Arguments <> '') then
  begin
    if (Pos('CHAINES', Arguments) > 0) then
    begin
      Origine := ReadTokenSt(Arguments);
      if trim(Arguments) <> '' then
      begin
        DateDebut := StrToDate(ReadTokenSt(Arguments));
        DateFin := StrToDate(ReadTokenSt(Arguments));
      end;
    end;
  end
  else Origine := 'MENU';
  //FIN PT4

  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  //PT3 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');

  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT7 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT7 }
  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PPU_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PPU_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PPU_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  CDd := THEdit(GetControl('XX_VARIABLEDEB'));
  CDf := THEdit(GetControl('XX_VARIABLEFIN'));
  if CDd <> nil then CDd.OnElipsisClick := DateElipsisclick;
  if CDf <> nil then CDf.OnElipsisClick := DateElipsisclick;
  if RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) = True then
  begin
    if CDd <> nil then CDd.text := DebPer;
    if CDf <> nil then CDf.text := FinPer;
  end;

  //Evenement ONCHANGE
  Check := TCheckBox(GetControl('CN1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN5'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CETAB'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CALPHA'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  {PT9 mise en commentaire
  ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
  if ChMonInv<>nil then begin ChMonInv.OnClick:=MonnaieInverse; PgMonnaieInv:=ChMonInv.checked; End;}

  Check := TCheckBox(GetControl('CKEURO'));
  if Check <> nil then Check.Checked := VH_Paie.PGTenueEuro;
  {$IFNDEF EAGLCLIENT}
  TFQRS1(Ecran).FListe.OnClick := ExportFiche; //PT1
  {$ENDIF}
  //DEB PT4 Affect critère standard
  if (origine = 'CHAINES') and (DateFin <> idate1900) then
  begin
    if CDd <> nil then CDd.text := DateToStr(DateDebut);
    if CDf <> nil then CDf.text := DateToStr(DateFin);
  end;
  //FIN PT4
  Check := TCheckBox(GetControl('CKPERIODE'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;

end;


procedure TOF_PGJOURPAIE_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGJOURPAIE_ETAT.OnUpdate;
var
  SQL, Temp, Tempo, Critere: string;
  Pages: TPageControl;
  x: integer;
  Cdd, CDf: THEdit;
  BulCompl : TCheckBox;
begin
  CDd := THEdit(GetControl('XX_VARIABLEDEB'));
  CDf := THEdit(GetControl('XX_VARIABLEFIN'));

  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);

  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  {PGCritere:=Critere; PORTAGECWAS Suppression de la variable globale
  !!on ne tient pas compte de la selection Libellé pour le chargement des tobs}
  SetControlText('STWHERE', Critere);

  if GetControlText('CALPHA') = 'X' then //DEB PT2
  begin
    if GetControlText('LIBELLE') <> '' then Critere := Critere + ' AND PPU_LIBELLE>="' + GetControlText('LIBELLE') + '" ';
    if GetControlText('LIBELLE_') <> '' then Critere := Critere + ' AND PPU_LIBELLE<="' + GetControlText('LIBELLE_') + '" ';
  end; //FIN PT2
  { DEB PT11 }
  BulCompl := TCheckBox(GetControl('BULCOMPL'));
  if Assigned(BulCompl) then
   IF BulCompl.State = cbUnchecked then
     Critere := Critere + ' AND PPU_BULCOMPL<>"X" '
  Else
   IF BulCompl.State = cbChecked then
     Critere := Critere + ' AND PPU_BULCOMPL="X" ';
  { FIN PT11 }
  {Utilisé pour test développement états chaînés
  If IsValidDate(CDd.text) then StDate:='WHERE PPU_DATEDEBUT>="'+UsDateTime(StrtoDate(CDd.text))+'" ' else StDate:='';
  If IsValidDate(CDf.text) then StDate:=StDate+'AND PPU_DATEFIN<="'+UsDateTime(StrtoDate(CDf.text))+'" ' else StDate:='';
  TFQRS1(Ecran).WhereSQL:=RendRequeteSQLEtat('JOURNALPAIE',AjoutChamp,StDate+Critere,'ORDER BY '+AjoutTri+' PPU_SALARIE');
  }
  if (IsValidDate(CDd.text)) and (IsValidDate(CDf.text)) then
  begin
    { DEB PT6 Mise en commentaire : Modification de le requête
      Chargement des cumuls depuis Paieencours et non plus histocumsal par fonction @
    if GetControlText('CKPERIODE')='X' then
      SQL:= 'SELECT PPU_SALARIE,PSA_LIBELLE,PSA_PRENOM,'+AjoutChamp+' '+
      'PSA_DATEENTREE,PSA_DATESORTIE,PPU_CBRUT,PPU_CBRUTFISCAL,'+
      'PPU_CNETIMPOSAB,PPU_CNETAPAYER,PPU_CCOUTSALARIE,PPU_CCOUTPATRON,PPU_CBASESS '+
      'FROM PAIEENCOURS '+
      'LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE '+
      'WHERE PPU_DATEDEBUT>="'+UsDateTime(StrtoDate(CDd.text))+'" '+
      'AND PPU_DATEFIN<="'+UsDateTime(StrtoDate(CDf.text))+'" '+critere+' '+
      'ORDER BY '+AjoutTri+' PPU_SALARIE'
    Else }
// PT10 18/06/2004 PH V_50 FQ 11164 Base SS pratiquée au lieu de base SS remplacement de PPU_CBASESS par PPU_CBASESSPRAT
    SQL := 'SELECT DISTINCT PPU_SALARIE,PSA_LIBELLE,PSA_PRENOM,' + AjoutChamp + ' ' +
      'PSA_DATEENTREE,PSA_DATESORTIE,' +
      'SUM(PPU_CHEURESTRAV) PPU_CHEURESTRAV,SUM(PPU_CBRUT) PPU_CBRUT,' +
      'SUM(PPU_CBRUTFISCAL) PPU_CBRUTFISCAL,SUM(PPU_CNETIMPOSAB) PPU_CNETIMPOSAB,' +
      'SUM(PPU_CNETAPAYER) PPU_CNETAPAYER,SUM(PPU_CCOUTSALARIE) PPU_CCOUTSALARIE,' +
      'SUM(PPU_CCOUTPATRON) PPU_CCOUTPATRON,SUM(PPU_CBASESSPRAT) PPU_CBASESSPRAT ' +
      'FROM PAIEENCOURS ' +
      'LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(StrtoDate(CDd.text)) + '" ' + //PT8
    'AND PPU_DATEFIN<="' + UsDateTime(StrtoDate(CDf.text)) + '" ' + critere + ' ' +
      'GROUP BY PPU_SALARIE,PSA_LIBELLE,PSA_PRENOM,' + AjoutChamp + ' PSA_DATEENTREE,PSA_DATESORTIE ' +
      'ORDER BY ' + AjoutTri + ' PPU_SALARIE';
    { FIN PT6 }
    TFQRS1(Ecran).WhereSQL := SQL;
  end
  else
    HShowMessage('5;Saisie :;Vous devez saisir une période de debut et de fin de paie!;W;O;O;O;;;', '', '');
end;


{PT9 Mise en commentaire
procedure TOF_PGJOURPAIE_ETAT.MonnaieInverse(Sender: TObject);
var
ChMonInv : TCheckBox;
begin
ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
if (ChMonInv<>nil)  then  PgMonnaieInv:=ChMonInv.checked;
PgTauxConvert:=RendTauxConvertion;
end;   }

procedure TOF_PGJOURPAIE_ETAT.ChangeLieuTravail(Sender: TObject);
var
  CEtab, CN1, CN2, CN3, CN4, CN5, Alpha, CL1, CL2, CL3, CL4: TCheckBox;
begin
  AjoutChamp := '';
  AjoutTri := '';
  BloqueChampLibre(Ecran);
  RecupChampRupture(Ecran);

  CEtab := TCheckBox(GetControl('CETAB'));
  if (CEtab <> nil) then
    if (CEtab.Checked = True) then
    begin
      AjoutChamp := 'PPU_ETABLISSEMENT, ';
      AjoutTri := 'PPU_ETABLISSEMENT, ';
      PgChampRupt := 'PPU_ETABLISSEMENT';
    end;
  CN1 := TCheckBox(GetControl('CN1'));
  CN2 := TCheckBox(GetControl('CN2'));
  CN3 := TCheckBox(GetControl('CN3'));
  CN4 := TCheckBox(GetControl('CN4'));
  CN5 := TCheckBox(GetControl('CN5'));
  Alpha := TCheckBox(GetControl('CALPHA'));

  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) then
  begin
    if (CN1.Checked = True) then
    begin
      AjoutChamp := 'PPU_TRAVAILN1, ';
      AjoutTri := 'PPU_TRAVAILN1, ';
      PgChampRupt := 'PPU_TRAVAILN1';
    end;
    if (CN2.Checked = True) then
    begin
      AjoutChamp := 'PPU_TRAVAILN2, ';
      AjoutTri := 'PPU_TRAVAILN2, ';
      PgChampRupt := 'PPU_TRAVAILN2';
    end;
    if (CN3.Checked = True) then
    begin
      AjoutChamp := 'PPU_TRAVAILN3, ';
      AjoutTri := 'PPU_TRAVAILN3, ';
      PgChampRupt := 'PPU_TRAVAILN3';
    end;
    if (CN4.Checked = True) then
    begin
      AjoutChamp := 'PPU_TRAVAILN4, ';
      AjoutTri := 'PPU_TRAVAILN4, ';
      PgChampRupt := 'PPU_TRAVAILN4';
    end;
    if (CN5.Checked = True) then
    begin
      AjoutChamp := 'PPU_CODESTAT, ';
      AjoutTri := 'PPU_CODESTAT, ';
      PgChampRupt := 'PPU_CODESTAT';
    end;
  end;

  CL1 := TCheckBox(GetControl('CL1'));
  CL2 := TCheckBox(GetControl('CL2'));
  CL3 := TCheckBox(GetControl('CL3'));
  CL4 := TCheckBox(GetControl('CL4'));
  if (CL1 <> nil) and (CL2 <> nil) and (CL3 <> nil) and (CL4 <> nil) then
  begin
    if (CL1.Checked = True) then
    begin
      AjoutChamp := 'PPU_LIBREPCMB1, ';
      AjoutTri := 'PPU_LIBREPCMB1, ';
      PgChampRupt := 'PPU_LIBREPCMB1';
    end;
    if (CL2.Checked = True) then
    begin
      AjoutChamp := 'PPU_LIBREPCMB2, ';
      AjoutTri := 'PPU_LIBREPCMB2, ';
      PgChampRupt := 'PPU_LIBREPCMB2';
    end;
    if (CL3.Checked = True) then
    begin
      AjoutChamp := 'PPU_LIBREPCMB3, ';
      AjoutTri := 'PPU_LIBREPCMB3, ';
      PgChampRupt := 'PPU_LIBREPCMB3';
    end;
    if (CL4.Checked = True) then
    begin
      AjoutChamp := 'PPU_LIBREPCMB4, ';
      AjoutTri := 'PPU_LIBREPCMB4, ';
      PgChampRupt := 'PPU_LIBREPCMB4';
    end;
  end;

  // TRI SUR LIBELLE SALARIE + CHAMPS RUPTURE    //Cas Combiné avec tri alphabétique
  if Alpha <> nil then
    if Alpha.Checked = True then
    begin
      AjoutTri := 'PSA_LIBELLE,';
    end;
  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) then
  begin
    if (Cetab.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_ETABLISSEMENT,PSA_LIBELLE,';
    if (CN1.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_TRAVAILN1,PSA_LIBELLE,';
    if (CN2.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_TRAVAILN2,PSA_LIBELLE,';
    if (CN3.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_TRAVAILN3,PSA_LIBELLE,';
    if (CN4.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_TRAVAILN4,PSA_LIBELLE,';
    if (CN5.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_CODESTAT,PSA_LIBELLE,';
  end;
  if (CL1 <> nil) and (CL2 <> nil) and (CL3 <> nil) and (CL4 <> nil) then
  begin
    if (CL1.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_LIBREPCMB1,PSA_LIBELLE,';
    if (CL2.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_LIBREPCMB2,PSA_LIBELLE,';
    if (CL3.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_LIBREPCMB3,PSA_LIBELLE,';
    if (CL4.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PPU_LIBREPCMB4,PSA_LIBELLE,';
  end;
  if TCheckBox(Sender).Name = 'CALPHA' then //PT2
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PPU_SALARIE', 'LIBELLE');
  {DEB PT5 Rupture période}
  if GetControlText('CKPERIODE') = 'X' then
  begin
    AjoutChamp := AjoutChamp + 'PPU_DATEDEBUT,PPU_DATEFIN,';
    AjoutTri := AjoutTri + 'PPU_DATEDEBUT,PPU_DATEFIN,';
    SetControlText('XX_RUPTURE2', 'PPU_DATEFIN');
  end
  else
    SetControlText('XX_RUPTURE2', '');
  {FIN PT5}
end;

procedure TOF_PGJOURPAIE_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;
//DEB PT1
procedure TOF_PGJOURPAIE_ETAT.ExportFiche(Sender: TObject);
begin
  {$IFNDEF EAGLCLIENT}
  if TFQRS1(Ecran).FListe.Checked = True then
  begin
    TFQRS1(Ecran).FEtat.Value := 'PJL';
    TFQRS1(Ecran).CodeEtat := 'PJL';
  end
  else
  begin
    TFQRS1(Ecran).FEtat.Value := 'PJP';
    TFQRS1(Ecran).CodeEtat := 'PJP';
  end;
  {$ENDIF}
end;
//FIN PT1
procedure TOF_PGJOURPAIE_ETAT.OnLoad;
begin
  inherited;
  //DEB PT4
  if (Origine = 'CHAINES') and (DateDebut > idate1900) then
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(DateDebut));
    SetControlText('XX_VARIABLEFIN', DateToStr(DateFin));
  end;
  //FIN PT4
end;

initialization
  registerclasses([TOF_PGJOURPAIE_ETAT]);
end.

