unit UTofEtatTarifTTC;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, UTOF, EntGc, UTob,
{$IFDEF EAGLCLIENT}
     eQRS1,
{$ELSE}
     QRS1, DBTables,
{$ENDIF}
     hmsgbox, M3FP, HCtrls, HEnt1, CalcOLEGescom, UtilArticle, UtilGc, Ent1;

type

  TOF_EtatTarifTTC = class(TOF)

    procedure OnArgument(StArgument: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
    procedure OnClickEtablissement;
    
  public
    NatureType, Regime, Devise : String;
  end;
  //Procedure init;
implementation

procedure TOF_EtatTarifTTC.OnClose;
begin
  VH_GC.TOBEdt.ClearDetail;
  initvariable;
end;

procedure TOF_EtatTarifTTC.OnUpdate;
begin
  SetControlText('REGIME', Regime) ;
  SetControlText('DEVISE', Devise);
  SetControlText('NATURETYPE', NatureType);
  VH_GC.TOBEdt.ClearDetail;
  initvariable;
end;

procedure TOF_EtatTarifTTC.OnArgument(StArgument: string);
var iCol, i: integer;
   Arguments, Critere, ChampMul, ValMul, TypeEtat: string;
begin
  Arguments := StArgument;
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      i := pos('=', Critere);
      if i <> 0 then
      begin
        ChampMul := copy(Critere, 1, i - 1);
        ValMul := copy(Critere, i + 1, length(Critere));
        if ChampMul = 'ETAT' then TypeEtat := ValMul;
        if ChampMul = 'TYPE' then NatureType := ValMul;
      end;
    end;
  until Critere = '';
  // Paramétrage des libellés des familles, stat. article et dimensions
  for iCol := 1 to 3 do ChangeLibre2('TGA_FAMILLENIV' + InttoStr(iCol), Ecran);
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
    for iCol := 4 to 8 do ChangeLibre2('TGA2_FAMILLENIV' + InttoStr(iCol), Ecran);
  if VH^.EtablisDefaut <> '' then
    THValComboBox(GetControl('ETABLISSEMENT')).Value := VH^.EtablisDefaut else
    if THValComboBox(GetControl('ETABLISSEMENT')).Values.Count > 0 then
    THValComboBox(GetControl('ETABLISSEMENT')).Value := THValComboBox(GetControl('ETABLISSEMENT')).Values[0];

  if NatureType = 'VTE' then REGIME := 'TTC' else REGIME := 'HT';
  DEVISE := V_PGI.DevisePivot ;
  if TypeEtat = 'ETATTARIF' then
  begin
    THLabel(GetControl('TGL_DATETARIF1')).Caption := TraduireMemoire('&Date du tarif');
    THLabel(GetControl('TGL_DATETARIF2')).Visible := False;
    THEdit(GetControl('DATETARIF2')).Visible := False;
    TFQRS1(Ecran).NatureEtat := 'TAR';
    TFQRS1(Ecran).CodeEtat := 'TAR';
  end else
  begin
    Ecran.Caption := TraduireMemoire('Etat comparatif des tarifs articles');
    UpdateCaption(Ecran);
  end;
end;

procedure TOF_EtatTarifTTC.OnClickEtablissement;
var etablissement: string;
  Q: TQuery;
begin
  etablissement := GetControlText('ETABLISSEMENT');
  if (etablissement <> '') then
  begin
    SetControlEnabled('FTARIFETABLIS', TRUE);
    SetControlEnabled('FTARIFETABLIS', FALSE);
  end;
  Q := OpenSQL('Select ET_DEVISE from ETABLISS where ET_ETABLISSEMENT="' + etablissement + '"', True);
  if not Q.EOF then DEVISE := Q.Fields[0].AsString ;
  Ferme(Q);
end;

procedure AGLOnClickEtablissement(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFQRS1) then TOTOF := TFQRS1(F).LaTOF
  else exit;
  if (TOTOF is TOF_EtatTarifTTC) then TOF_EtatTarifTTC(TOTOF).OnClickEtablissement;
end;
initialization
  registerclasses([TOF_EtatTarifTTC]);
  RegisterAglProc('OnClickEtablissement', TRUE, 0, AGLOnClickEtablissement);
end.
