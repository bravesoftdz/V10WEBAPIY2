{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 18/04/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : LIBRE;PAIE
*****************************************************************
PT1 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono}
unit UTOFPGEtatLibre;
interface
uses  Controls, Classes,sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  UtileAGL, MaineAgl, eQRS1,
  {$ELSE}
   QRS1,
  {$ENDIF}
  HCtrls, HEnt1,  UTOF,

  HQry, ParamSoc;


type
  TOF_PGLIBRE_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure Change(Sender: TObject);
    procedure InitJourMois(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
  end;

implementation
uses PGEditOutils2, PgOutils, PgOutils2, Entpaie, PgEdtEtat;
{ TOF_PGLIBRE_ETAT }

procedure TOF_PGLIBRE_ETAT.Change(Sender: TObject);
begin

end;

procedure TOF_PGLIBRE_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGLIBRE_ETAT.InitJourMois(Sender: TObject);
var
  Mois, Mois2, Annee, FdM, DDM, AnneeOk: string;
  JJ, MM, YY: WORD;
begin
  Mois := GetControlText('CBMOIS');
  Annee := GetControlText('CBANNEE');
  if (mois = '') or (annee = '') then exit;
  ControlMoisAnneeExer(Mois, RechDom('PGANNEESOCIALE', Annee, FALSE), AnneeOk);
  Annee := AnneeOk;
  DDM := DateToStr(EncodeDate(StrToInt(Annee), StrToInt(Mois), 1));
  FdM := DateToStr(FinDeMois(PlusMois(StrToDate(DDM), 1)));
  SetControlText('DATEDEBUT', DDM);
  SetControlText('DATEFIN', FdM);
  Decodedate(StrToDate(DDM), YY, MM, JJ);
  SetControlText('ANNEEMOIS1', IntToStr(YY));
  Decodedate(StrToDate(FDM), YY, MM, JJ);
  SetControlText('ANNEEMOIS2', IntToStr(YY));
  if Mois <> '12' then Mois2 := IntToStr(StrToInt(Mois) + 1) else Mois2 := '01';
  if Length(Mois2) = 1 then Mois2 := '0' + Mois2;
  SetControlText('CBMOIS_', Mois2);
end;

procedure TOF_PGLIBRE_ETAT.OnArgument(Arguments: string);
var
  Min, Max, Mois, ExerPerEncours, DebPer, FinPer: string;
  YY, MM, JJ: Word;
  Combo: THValComboBox;
  Defaut: ThEdit;
  Zone: Tcontrol;
begin
  inherited;
  Combo := THValComboBox(GetControl('CBMOIS'));
  if Combo <> nil then Combo.OnChange := InitJourMois;
  Combo := THValComboBox(GetControl('CBANNEE'));
  if Combo <> nil then Combo.OnChange := InitJourMois;
  Zone := ThValComboBox(getcontrol('PSA_CALENDRIER'));
  InitialiseCombo(Zone);
  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PSA_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  //PT1 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  DecodeDate(StrToDate(DebPer), YY, MM, JJ);
  Mois := IntToStr(MM);
  if Length(Mois) = 1 then Mois := '0' + Mois;
  SetControlproperty('CBMOIS', 'Value', Mois);
  SetControlproperty('CBANNEE', 'Value', ExerPerEncours);
  if Ecran <> nil then PGCalendrierEtat_Ecran := Ecran;
end;

procedure TOF_PGLIBRE_ETAT.OnUpdate;
var
  SQL, Temp, Tempo, Critere: string;
  DD, DF: TDateTime;
  Pages: TPageControl;
  x: integer;
begin
  inherited;
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  DD := StrToDate(GetControlText('DATEDEBUT'));
  DF := StrToDate(GetControlText('DATEFIN'));
  if (DD > 2) and (DF > 2) then
    SQL := 'SELECT DISTINCT PSA_SALARIE,PSA_LIBELLE, PSA_ETABLISSEMENT,PSA_STANDCALEND,' +
      'PSA_CALENDRIER,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_JOURHEURE ' +
      'FROM SALARIES ' +
      'WHERE PSA_DATEENTREE<="' + USDateTime(DF) + '" ' +
      'AND (PSA_DATESORTIE>="' + USDateTime(DD) + '" OR PSA_DATESORTIE="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) ' +
      '' + critere + ' ' +
      'ORDER BY PSA_SALARIE ';
  TFQRS1(Ecran).WhereSQL := SQL;
end;

initialization
  registerclasses([TOF_PGLIBRE_ETAT]);
end.

