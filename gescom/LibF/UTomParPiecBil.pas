{***********UNITE*************************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/03/2001
Modifié le ... : 05/08/2004
Description .. : Paramétrage des pièces et billets
Mots clefs ... : FO
*****************************************************************}
unit UTomParPiecBil;

interface
uses
  SysUtils, Forms, Classes, controls, StdCtrls,
  {$IFDEF EAGLCLIENT}
  eFichList,
  {$ELSE}
  FichList, db,
  {$ENDIF}
  M3FP, HCtrls, HEnt1, UTOB, UTOM;

type
  TOM_ParPiecBil = class(TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( stArgument: string )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    ModeExec: string; // mode BO ou FO
    procedure ActiveBoutons;
    procedure FORefreshDevise(stRange: string);
    procedure DefMasqueDevise(CodeDevise: string);
  end;

implementation

const
  // Libellés des messages
  TexteMessage: array[1..1] of string = (
    {1} 'Vous devez renseigner une valeur correcte !'
    );

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOM_ParPiecBil.OnArgument(stArgument: string);
var
  ChampMul, ValMul, Critere: string;
  Ind: integer;
begin
  inherited;
  repeat
    Critere := Trim(ReadTokenSt(stArgument));
    if Critere <> '' then
    begin
      Ind := pos('=', Critere);
      if Ind <> 0 then
      begin
        ChampMul := copy(Critere, 1, Ind - 1);
        ValMul := copy(Critere, Ind + 1, length(Critere));
        if ChampMul = 'MODE' then ModeExec := ValMul;
        if ChampMul = 'DEVISE' then SetControlText('DEVISE', ValMul);
      end;
    end;
  until Critere = '';
  ActiveBoutons;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : OnChangeField
Mots clefs ... :
*****************************************************************}

procedure TOM_ParPiecBil.OnChangeField(F: TField);
begin
  inherited;
  if F.FieldName = 'GPI_DEVISE' then DefMasqueDevise(GetField('GPI_DEVISE'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : OnCancelRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_ParPiecBil.OnCancelRecord;
var
  sRange: string;
  Ctrl: TControl;
begin
  inherited;
  if (Ecran <> nil) and (Ecran is TFFicheListe) then
  begin
    sRange := GetControlText('DEVISE');
    Ctrl := GetControl('RB_PIECES');
    if (Ctrl <> nil) and (Ctrl is TRadioButton) and TRadioButton(Ctrl).Checked then
      sRange := sRange + ';PIE;'
    else
      sRange := sRange + ';BIL;';
    FORefreshDevise(sRange);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnNewRecord
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnNewRecord ;
begin
  Inherited ;
  if (ModeExec = 'FO') and (Ecran <> nil) then
  begin
    if Ecran is TFFicheListe then TFFicheListe(Ecran).BDefaire.Click;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnDeleteRecord
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnUpdateRecord
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnUpdateRecord ;
begin
  Inherited ;
  if GetField('GPI_VALEUR') <= 0 then
  begin
    SetFocusControl('GPI_VALEUR');
    LastError := 1;
    LastErrorMsg := TexteMessage[1];
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnAfterUpdateRecord
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnAfterDeleteRecord
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnLoadRecord
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnLoadRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/08/2004
Modifié le ... : 05/08/2004
Description .. : OnClose
Mots clefs ... : 
*****************************************************************}

procedure TOM_PARPIECBIL.OnClose ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : constitue le masque de saisie de la valeur en fonction de la
Suite ........ : devise
Mots clefs ... :
*****************************************************************}

procedure TOM_ParPiecBil.DefMasqueDevise(CodeDevise: string);
var
  Stg: string;
  NbDec: integer;
begin
  Stg := GetColonneSQL('DEVISE', 'D_DECIMALE', 'D_DEVISE="'+ CodeDevise +'"');
  if IsNumeric(Stg) then NbDec := StrToInt(Stg) else NbDec := V_PGI.OkDecV;
  SetControlProperty('GPI_VALEUR', 'DisplayFormat', StrfMask(NbDec, '', True));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : E. VIENOT
Créé le ...... : 15/07/2004
Modifié le ... : 15/07/2004
Description .. : Inactive les boutons selon le mode d'exécution
Mots clefs ... : 
*****************************************************************}

procedure TOM_ParPiecBil.ActiveBoutons;
begin
  if ModeExec = 'FO' then // mode consultation
  begin
    SetControlEnabled('GPI_PIECEBILLET', False);
    SetControlEnabled('GPI_LIBELLE', False);
    SetControlEnabled('GPI_ABREGE', False);
    SetControlEnabled('GPI_VALEUR', False);
    SetControlEnabled('bInsert', True);
    SetControlVisible('bInsert', False);
    SetControlVisible('bDelete', False);        
    SetControlVisible('bValider', False);
    SetControlVisible('bDefaire', False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : maj du range si changement de devise ou du type de
Suite ........ : l'espèce
Mots clefs ... :
*****************************************************************}

procedure TOM_ParPiecBil.FORefreshDevise(stRange: string);
var
  FF: TFFicheListe;
begin
  FF := TFFicheListe(TForm(Ecran));
  FF.FRange := stRange;
  {$IFDEF EAGLCLIENT}
  FF.ReLoadDB;
  FF.ChargeEnreg;
  {$ELSE}
  SetLeRange(FF.Ta, FF.FRange);
  {$ENDIF}
  ActiveBoutons;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FORefreshDevise = maj du range si changement de devise ou
Suite ........ : du type de l'espèce depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Suite ........ :  - Parms[1] = Range
Mots clefs ... : FO
*****************************************************************}

procedure AGLRefreshDevise(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_ParPiecBil) then TOM_ParPiecBil(OM).FORefreshDevise(Parms[1]) else exit;
end;

initialization
  registerclasses([TOM_ParPiecBil]);
  RegisterAglProc('FORefreshDevise', TRUE, 1, AGLRefreshDevise);
end.
