{***********UNITE*************************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... : 30/03/2007
Description .. : TOF de la fiche IMPRDUCS (QRS1)
Suite ........ : Lancement état DUC (Ducs), DUN (DUCS Néanr), PVU 
Suite ........ : (VLU annexe à la Ducs Assedic)
Suite ........ : Appelé des menu 42302, 42303 et du menu des états 
Suite ........ : chaînes
Suite ........ : 
Suite ........ : C'est sur cette fiche qu'on pourrait traiter la progammation 
Suite ........ : du Process Server.
Mots clefs ... : PAIE; DUCS; ETATSCHAINES;
*****************************************************************}
unit UtofPGDucs;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}

  Classes,
  Hent1,
  HCtrls,
{$IFDEF EAGLCLIENT}
  eQRS1,
{$ELSE}
  QRS1,
{$ENDIF}
  PgDucsOutils,
  PgEdtEtat,
  sysutils,
  PGVLU,
  UTOF,
  UTOFPGEtats;


type
  TOF_PGIMPRDUCS = class(TOF_PGEtats)
  private
    Origine: string;
    DD, DF: TDateTime;
    ND, NF, ED, EF  : string;
    DebPer, FinPer, PeriodEtat, NatDeb, NatFin, EtabDeb, EtabFin, CodeEtat :string;
    Neant : boolean;

    public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnLoad; override;

end;


implementation
procedure TOF_PGIMPRDUCS.OnArgument(Arguments: string);
begin
  inherited;
  Neant := false;
  DD := idate1900; DF := idate1900;
  if (Arguments <> '') then
  begin
    if (Pos('CHAINES', Arguments) > 0) then
    // Etats Chaînés
    begin
      Origine := ReadTokenSt(Arguments);
      if trim(Arguments) <> '' then
      // récupération critères d'édition
      begin
        DD := StrToDate(ReadTokenSt(Arguments));
        DF := StrToDate(ReadTokenSt(Arguments));
        ND := ReadTokenSt(Arguments);
        NF := ReadTokenSt(Arguments);
        ED := ReadTokenSt(Arguments);
        EF := ReadTokenSt(Arguments);
        CodeEtat := ReadTokenSt(Arguments);
      end;
    end
    else
    // appel menu 42302 ou 42303 
    begin
      Origine :=ReadTokenSt(Arguments);
      CodeEtat := ReadTokenSt(Arguments);
    end;

    if ((Origine = 'CgiPaie') or (Origine ='CHAINES')) and (CodeEtat = 'DUN') then
        Neant := true;

    if (Origine = 'PVU') then
    begin
      CodeEtat := 'PVU';
      SetControlVisible('NATUREORG', false);
      SetControlVisible('NATUREORG_', false);
      SetControlVisible('PDU_ETABLISSEMENT', false);
      SetControlVisible('PDU_ETABLISSEMENT_', false);
      SetControlVisible('TPDU_ETABLISSEMENT', false);
      SetControlVisible('TPDU_ETABLISSEMENT_', false);
      SetControlVisible('TPDU_ORGANISME', false);
      SetControlVisible('TPDU_ORGANISME_', false);
      TFQRS1(Ecran).Caption := 'Impression VLU';
      UpdateCaption(TFQRS1(Ecran));
    end;
    if (CodeEtat <> '') then
    begin
      SetControlVisible('NEANT', false);
      SetControlEnabled('NEANT', false);
    end;
  end;
end;
procedure TOF_PGIMPRDUCS.OnUpdate;
var
  StSql   : string;
begin
  StSql := '';

  TFQRS1(Ecran).FCodeEtat := CodeEtat;

  if (CodeEtat = 'PVU') then
  // impression VLU
  begin
    CalculPeriode (PeriodEtat, DebPer, FinPer);
    if (origine = 'CHAINES') or (origine = 'PVU') then
      // état chaîné ou Impression VLU (menu 42303)
      EditVLU(StrToDate(Debper), StrToDate(FinPer), '002', Pages, True, false,StSQL)
    else
      // A voir qd ?
      EditVLU(StrToDate(Debper), StrToDate(FinPer), '002', Pages, False, false,StSQL)
  end
  else
    // Impression DUCS (menu 42302 ou Etat Chaîné
    DucsEdit (PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin, Neant, StSql);

  FunctPGDebEditDucs(); {raz n° page}
  TFQRS1(Ecran).WhereSQL := StSql;
  FunctPGFinEditDucs(); {raz AncNumPagSys}
end;

procedure TOF_PGIMPRDUCS.OnLoad;
begin
  inherited;

  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    SetControlText('PDU_DATEDEBUT', DateToStr(DD));
    SetControlText('PDU_DATEFIN', DateToStr(DF));
    SetControlText('NATUREORG', ND);
    SetControlText('NATUREORG_', NF);
    SetControlText('PDU_ETABLISSEMENT', ED);
    SetControlText('PDU_ETABLISSEMENT_', EF);
  end;

  DebPer := GetControlText('PDU_DATEDEBUT');
  FinPer := GetControlText('PDU_DATEFIN');
  NatDeb := GetControlText('NATUREORG');
  NatFin := GetControlText('NATUREORG_');
  EtabDeb := GetControlText('PDU_ETABLISSEMENT');
  EtabFin := GetControlText('PDU_ETABLISSEMENT_');

  if (origine <> 'CHAINES') and (origine <> 'CgiPaie') and (CodeEtat <> 'PVU') then
  // Impression DUCS (menu 42302)
  begin
    if (GetControlText('NEANT') = 'X') then
    // Ducs Néant
    begin
      CodeEtat := 'DUN';
      Neant := true;
    end
    else
    // Ducs non Néant
    begin
      CodeEtat := 'DUC';
      Neant := false;
    end;
  end;
end;

initialization
  registerclasses([TOF_PGIMPRDUCS]);
end.

 