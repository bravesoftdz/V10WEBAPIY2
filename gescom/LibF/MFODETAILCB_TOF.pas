{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 28/11/2003
Description .. : Source TOF des FICHES :
Suite ........ :  * MFODETAILCB - Saisie du détail de la carte bancaire ou
Suite ........ : du chèque
Suite ........ :  * MFOCBFROMKB - Lecture d'une carte depuis la piste du
Suite ........ : clavier
Suite ........ :
Mots clefs ... : TOF;MFODETAILCB;MFOCBFROMKB
*****************************************************************}
Unit MFODETAILCB_TOF;

Interface

Uses
  Controls, Classes, sysutils, stdctrls,
{$IFDEF EAGLCLIENT}
  Maineagl,
{$else}
  Fe_Main, 
{$ENDIF}
  Vierge, HCtrls, HPanel, HEnt1, UTOB, UTOF, ParamSoc, LicUtil, AGLInit;

Type
  TOF_MFODETAILCB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Erreur: boolean;
    NomTiers: string;
    function CopieNumCB2Ctrl: boolean;
    procedure MetEnErreur(NomChamp: string; NoMsg: integer);
    procedure BrancheOnExit(NomCtrl: string; FctOnExit: TNotifyEvent);
    procedure GPE_CBINTERNETExit(Sender: TObject);
    procedure GPE_DATEEXPIREExit(Sender: TObject);
  end ;

Type
  TOF_MFOCBFROMKB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Erreur: boolean;
    SansClient: boolean;
  end ;

function FOLanceFicheLectureCBFromKB(Range, Lequel, Argument: string): string;
function FODecodeLectureCBFromKB(TOBEches, TOBLgEch, TOBMdp: TOB; Raz: boolean; ValeurLue: string): boolean;
function FOLanceFicheDetailCB(Range, Lequel, Argument: string): string;
function FOSaisieDetailCB(TOBEches, TOBLgEch, TOBMdp: TOB; NomClient: string; DepuisTPE: boolean = False): boolean;
function FOVerifDetailCB(TOBEch, TOBMdp: TOB): integer ;

Implementation

Uses
  UtilCB, UtofGCAcomptes, SaisUtil, AGLInitGC, FODefi, FOUtil;

const
  // Libellés des messages
  TexteMessage: array[1..5] of string = (
    {1}'Le numéro de carte n''est pas correct',
    {2}'La date d''expiration n''est pas correcte ou la carte est expirée',
    {3}'Le type de la carte n''est pas cohérent avec son numéro',
    {4}'Le numéro d''autorisation n''est pas renseigné',
    {5}'Le numéro de chèque n''est pas renseigné'
    );

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/11/2003
Modifié le ... : 28/11/2003
Description .. : Lancement de la fiche de lecture de la piste magnétique 
Suite ........ : d'une carte bancaire depuis un lecteur intégré dans le
Suite ........ : clavier
Mots clefs ... : 
*****************************************************************}

function FOLanceFicheLectureCBFromKB(Range, Lequel, Argument: string): string;
begin
  Result := AGLLanceFiche('MFO', 'MFOCBFROMKB', Range, Lequel, Argument);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : Lancement de la fiche de saisie du détail d'une carte
Suite ........ : bancaire ou d'un chèque
Mots clefs ... :
*****************************************************************}

function FOLanceFicheDetailCB(Range, Lequel, Argument: string): string;
begin
  Result := AGLLanceFiche('MFO', 'MFODETAILCB', Range, Lequel, Argument);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/12/2003
Modifié le ... : 15/12/2003
Description .. : Saisie du détail d'une carte bancaire ou d'un chèque
Mots clefs ... :
*****************************************************************}

function FOSaisieDetailCB(TOBEches, TOBLgEch, TOBMdp: TOB; NomClient: string; DepuisTPE: boolean = False): boolean;
var
  InfoCompl, NumAutor, CodeMode, TypeMode, sArgs, sLib: string;
  Montant: double;
  DelCBKB: boolean;
begin
  Result := True;
  CodeMode := TOBLgEch.GetValue('GPE_MODEPAIE');
  Montant := TOBLgEch.GetValue('GPE_MONTANTENCAIS');
  TypeMode := TOBMdp.GetValue('MP_TYPEMODEPAIE');
  if DepuisTPE then
  begin
    InfoCompl := 'X';
    NumAutor := '-';
  end else
  begin
    InfoCompl := TOBMdp.GetValue('MP_AVECINFOCOMPL');
    NumAutor := TOBMdp.GetValue('MP_AVECNUMAUTOR');
  end;

  if (InfoCompl = 'X') or (NumAutor= 'X') then
  begin
    // Saisie des informations complémentaires
    sArgs := 'GPE_MODEPAIE=' + CodeMode + ';GPE_MONTANTENCAIS=' + StrSP(Montant)
      + ';MP_TYPEMODEPAIE=' + TypeMode
      + ';MP_AVECINFOCOMPL=' + InfoCompl
      + ';MP_AVECNUMAUTOR=' + NumAutor
      + ';MP_COPIECBDANSCTRL=' + TOBMdp.GetValue('MP_COPIECBDANSCTRL')
      + ';MP_AFFICHNUMCBUS=' + TOBMdp.GetValue('MP_AFFICHNUMCBUS');

    if FODecodeLectureCBFromKB(TOBEches, TOBLgEch, TOBMdp, False, '') then
    begin
      sLib := TOBLgEch.GetValue('GPE_CBLIBELLE');
      DelCBKB := True;
    end else
    begin
      sLib := '';
      DelCBKB := False;
    end;

    if sLib = '' then sLib := NomClient;
    sArgs := sArgs + ';GPE_CBLIBELLE=' + sLib;

    TheTob := TOBLgEch;
    if FOLanceFicheDetailCB('', '', sArgs) = 'OK' then
    begin
      if DelCBKB then TOBEches.DelChampSup('LECTURECBFROMKB', False);
    end else
      Result := False;
    TheTob := nil;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/12/2003
Modifié le ... : 15/12/2003
Description .. : Decode les informations lues sur une carte depuis la piste 
Suite ........ : du clavier
Mots clefs ... : 
*****************************************************************}

function FODecodeLectureCBFromKB(TOBEches, TOBLgEch, TOBMdp: TOB; Raz: boolean; ValeurLue: string): boolean;
var
  Stg, NumCB: string;
begin
  Result := False;
  if (not Assigned(TOBLgEch)) or (not Assigned(TOBMdp)) then Exit;

  if (Assigned(TOBEches)) and (TOBEches.FieldExists('LECTURECBFROMKB')) then
  begin
    Stg := TOBEches.GetValue('LECTURECBFROMKB');
    if Raz then TOBEches.DelChampSup('LECTURECBFROMKB', False);
  end else
    Stg := ValeurLue;

  if Stg <> '' then
  begin
    Result := True;
    NumCB := Trim(ReadTokenST(Stg));
    TOBLgEch.PutValue('GPE_CBINTERNET', CrypteNoCarteCB(NumCB));
    TOBLgEch.PutValue('GPE_TYPECARTE', Trim(ReadTokenST(Stg)));
    TOBLgEch.PutValue('GPE_DATEEXPIRE', Trim(ReadTokenST(Stg)));
    TOBLgEch.PutValue('GPE_CBLIBELLE', Trim(ReadTokenST(Stg)));
    // recopie du n° de carte dans le n° de contrôle
    if TOBMdp.GetValue('MP_COPIECBDANSCTRL') = 'X' then
    begin
      if TOBMdp.GetValue('MP_AFFICHNUMCBUS') = 'X' then
        Stg := Copy(NumCB, (Length(NumCB) -3), 4)  // recopie des 4 derniers caractères (format US)
      else
        Stg := Copy(NumCB, 7, 9);                  // recopie du 7ème au 15ème caractère (format français)
      TOBLgEch.PutValue('GPE_CBNUMCTRL', Stg);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

function FOVerifDetailCB(TOBEch, TOBMdp: TOB): integer ;
var
  TypeMdp, sNumCarte: string;
  OkInfo, OkAuto: boolean;
begin
  Result := 0;
  if (TOBEch = nil) or (TOBMdp = nil) then Exit;

  TypeMdp := TOBMdp.GetValue('MP_TYPEMODEPAIE');
  OkAuto := (TOBMdp.GetValue('MP_AVECNUMAUTOR') = 'X');
  OkInfo := (TOBMdp.GetValue('MP_AVECINFOCOMPL') = 'X');

  if not GetParamSoc('SO_GCCBFACULTATIF') then
  begin
    if OkInfo then
    begin
      if TypeMdp = TYPEPAIECB then
      begin
        sNumCarte := DeCrypteNoCarteCB(TOBEch.GetValue('GPE_CBINTERNET'), False, True, True);
        if not VerifNoCarteCB(sNumCarte) then
          Result := 1
        else if not VerifDateExpireCB(TOBEch.GetValue('GPE_DATEEXPIRE')) then
          Result := 2
        else if not GoodCBNumberCode(sNumCarte, TOBEch.GetValue('GPE_TYPECARTE')) then
          Result := 3;
      end else
      if (TypeMdp = TYPEPAIECHEQUE) or (TypeMdp = TYPEPAIECHQDIFF) then
      begin
        if Trim(TOBEch.GetValue('GPE_NUMCHEQUE')) = '' then
          Result := 5;
      end;
    end;

    if (Result = 0) and (OkAuto) then
    begin
      if Trim(TOBEch.GetValue('GPE_CBNUMAUTOR')) = '' then
        Result := 4;
    end;
  end;
end ;

{==============================================================================================}
{================================== TOF_MFODETAILCB ===========================================}
{==============================================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/11/2003
Modifié le ... : 21/11/2003
Description .. : Active la recopie des 4 derniers caractères du n° de carte
Suite ........ : dans le n° de contrôle
Mots clefs ... :
*****************************************************************}

function TOF_MFODETAILCB.CopieNumCB2Ctrl: boolean;
begin
  Result := (GetCheckBoxState('MP_COPIECBDANSCTRL') = cbChecked);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : met un champ en erreur et affiche le message associé
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.MetEnErreur(NomChamp: string; NoMsg: integer);
begin
  SetFocusControl(NomChamp);
  LastError := NoMsg;
  Erreur := True;
  if (NoMsg >= Low(TexteMessage)) and (NoMsg <= High(TexteMessage)) then
    LastErrorMsg := TexteMessage[NoMsg];
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : Branche un gestionnaire d'événement OnClick sur un
Suite ........ : contrôle
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.BrancheOnExit(NomCtrl: string; FctOnExit: TNotifyEvent);
var
  Ctrl: TControl;
begin
  Ctrl := GetControl(NomCtrl);
  if Ctrl = nil then Exit;

  if Ctrl is THEdit then
  begin
    THEdit(Ctrl).OnExit := FctOnExit;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : Changement du n° de carte bancaire
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.GPE_CBINTERNETExit(Sender: TObject);
var
  NumCB, Stg: string;
begin
  NumCB := GetControlText('GPE_CBINTERNET');
  // recopie du n° de carte dans le n° de contrôle
  if CopieNumCB2Ctrl then
  begin
    if GetCheckBoxState('MP_AFFICHNUMCBUS') = cbChecked then
      Stg := Copy(NumCB, (Length(NumCB) -3), 4)  // recopie des 4 derniers caractères (format US)
    else
      Stg := Copy(NumCB, 7, 9);                  // recopie du 7ème au 15ème caractère (format français)
    SetControlText('GPE_CBNUMCTRL', Stg);
  end;
  // Recherche du type de carte en fonction de numéro
  Stg := CodeCardType(NumCB);
  SetControlText('GPE_TYPECARTE', Stg);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : Changement de la date d'expiration de la carte
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.GPE_DATEEXPIREExit(Sender: TObject);
var
  DateExpire, Annee: string;
begin
  DateExpire := GetControlText('GPE_DATEEXPIRE');
  Annee := Trim(Copy(DateExpire, 4, 4));
  if Length (Annee) = 2 then
  begin
    // cas de l'année saisie sur 2 caractères
    Insert('20', DateExpire, 4);
    SetControlText('GPE_DATEEXPIRE', DateExpire);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnUpdate ;
var
  TypeMdp, Stg: string;
  OkInfo, OkAuto: boolean;
begin
  Inherited ;
  //if Action = taConsult then Exit;
  LastErrorMsg := '';
  LastError := 0;
  Erreur := False;
  TypeMdp := GetControlText('MP_TYPEMODEPAIE');
  OkAuto := (GetCheckBoxState('MP_AVECNUMAUTOR') = cbChecked);
  OkInfo := (GetCheckBoxState('MP_AVECINFOCOMPL') = cbChecked);

  if not GetParamSoc('SO_GCCBFACULTATIF') then
  begin
    if OkInfo then
    begin
      if TypeMdp = TYPEPAIECB then
      begin
        if not VerifNoCarteCB(GetControlText('GPE_CBINTERNET')) then
          MetEnErreur('GPE_CBINTERNET', 1)
        else if not VerifDateExpireCB(GetControlText('GPE_DATEEXPIRE')) then
          MetEnErreur('GPE_DATEEXPIRE', 2)
        else if not GoodCBNumberCode(GetControlText('GPE_CBINTERNET'), GetControlText('GPE_TYPECARTE')) then
          MetEnErreur('GPE_TYPECARTE', 3);
      end else
      if (TypeMdp = TYPEPAIECHEQUE) or (TypeMdp = TYPEPAIECHQDIFF) then
      begin
        if Trim(GetControlText('GPE_NUMCHEQUE')) = '' then
          MetEnErreur('GPE_NUMCHEQUE', 5);
      end;
    end;

    if (not Erreur) and (OkAuto) then
    begin
      if Trim(GetControlText('GPE_CBNUMAUTOR')) = '' then
        MetEnErreur('GPE_CBNUMAUTOR', 4);
    end;
  end;

  if Erreur then
  begin
    TFVierge(Ecran).Retour := '';
  end else
  begin
    TFVierge(Ecran).Retour := 'OK';
    //LaTOB.GetEcran(Ecran);
    LaTOB.PutValue('GPE_CBINTERNET', '');
    LaTOB.PutValue('GPE_CBNUMCTRL', '');
    LaTOB.PutValue('GPE_DATEEXPIRE', '');
    LaTOB.PutValue('GPE_TYPECARTE', '');
    LaTOB.PutValue('GPE_NUMCHEQUE', '');
    LaTOB.PutValue('GPE_CBNUMAUTOR', '');
    LaTOB.PutValue('GPE_CBLIBELLE', '');
    if OkInfo then
    begin
      if (TypeMdp = TYPEPAIECHEQUE) or (TypeMdp = TYPEPAIECHQDIFF) then
      begin
        LaTOB.PutValue('GPE_NUMCHEQUE', GetControlText('GPE_NUMCHEQUE'));
      end else
      if TypeMdp = TYPEPAIECB then
      begin
        LaTOB.PutValue('GPE_CBINTERNET', CrypteNoCarteCB(GetControlText('GPE_CBINTERNET')));
        LaTOB.PutValue('GPE_CBNUMCTRL', GetControlText('GPE_CBNUMCTRL'));
        LaTOB.PutValue('GPE_DATEEXPIRE', GetControlText('GPE_DATEEXPIRE'));
        LaTOB.PutValue('GPE_TYPECARTE', GetControlText('GPE_TYPECARTE'));
        Stg := GetControlText('GPE_CBLIBELLE');
        if Stg <> NomTiers then LaTOB.PutValue('GPE_CBLIBELLE', Stg);
      end;
    end;
    if OkAuto then LaTOB.PutValue('GPE_CBNUMAUTOR', GetControlText('GPE_CBNUMAUTOR'));
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnLoad ;
var
  TypeMdp, Stg: string;
  Ctrl: TControl;
  OkInfo, OkAuto, Visu: boolean;
begin
  Inherited ;
  TypeMdp := GetControlText('MP_TYPEMODEPAIE');
  OkAuto := (GetCheckBoxState('MP_AVECNUMAUTOR') = cbChecked);
  Visu := OkAuto;
  SetControlVisible('PANELBAS', Visu);

  OkInfo := (GetCheckBoxState('MP_AVECINFOCOMPL') = cbChecked);
  Visu := ((TypeMdp = TYPEPAIECB) and (OkInfo));
  SetControlVisible('PANELNUMCARTE', Visu);
  Visu := (((TypeMdp = TYPEPAIECHEQUE) or (TypeMdp = TYPEPAIECHQDIFF)) and (OkInfo));
  SetControlVisible('PANELNUMCHEQUE', Visu);

  if Assigned(Ecran) then
  begin
    if (TypeMdp = TYPEPAIECHEQUE) or (TypeMdp = TYPEPAIECHQDIFF) then
    begin
      Ecran.Caption := TraduireMemoire('Détail du chèque');
      UpdateCaption(Ecran);
    end else
    if TypeMdp = TYPEPAIECB then
    begin
      Ecran.Caption := TraduireMemoire('Détail de la carte bancaire');
      UpdateCaption(Ecran);
    end;

    if (not OkInfo) or ((TypeMdp <> TYPEPAIECHEQUE) and (TypeMdp <> TYPEPAIECHQDIFF)) then
    begin
      Ctrl := GetControl('PANELNUMCHEQUE');
      if (Assigned(Ctrl)) and (Ctrl is THPanel) then
        Ecran.Height := Ecran.Height - THPanel(Ctrl).Height;
    end;
    if (not OkInfo) or (TypeMdp <> TYPEPAIECB) then
    begin
      Ctrl := GetControl('PANELNUMCARTE');
      if (Assigned(Ctrl)) and (Ctrl is THPanel) then
        Ecran.Height := Ecran.Height - THPanel(Ctrl).Height;
    end;
  end;

  //LaTOB.PutEcran(Ecran);
  SetControlText('GPE_CBNUMAUTOR', LaTOB.GetValue('GPE_CBNUMAUTOR'));
  SetControlText('GPE_NUMCHEQUE', LaTOB.GetValue('GPE_NUMCHEQUE'));
  SetControlText('GPE_CBINTERNET', DeCrypteNoCarteCB(LaTOB.GetValue('GPE_CBINTERNET'), False, True, True));
  SetControlText('GPE_CBNUMCTRL', LaTOB.GetValue('GPE_CBNUMCTRL'));
  SetControlText('GPE_DATEEXPIRE', LaTOB.GetValue('GPE_DATEEXPIRE'));
  SetControlText('GPE_TYPECARTE', LaTOB.GetValue('GPE_TYPECARTE'));
  Stg := LaTOB.GetValue('GPE_CBLIBELLE');
  if Stg = '' then Stg := NomTiers;
  SetControlText('GPE_CBLIBELLE', Stg); 
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnArgument (S : String ) ;
var
  Stg, NomChamp, ValeurChamp: string;
  Ind: integer;
begin
  Inherited ;
  TFVierge(Ecran).Retour := '';
  NomTiers := '';
  Stg := S;
  repeat
    Stg := UpperCase(Trim(ReadTokenSt(S)));
    if Stg <> '' then
    begin
      Ind := Pos('=', Stg);
      if Ind <> 0 then
      begin
        NomChamp := Copy(Stg, 1, (Ind - 1));
        ValeurChamp := Trim(Copy(Stg, (Ind + 1), MaxInt));
        SetControlText(NomChamp, ValeurChamp);
        if NomChamp = 'GPE_CBLIBELLE' then NomTiers := ValeurChamp;
      end;
    end;
  until S = '';

  BrancheOnExit('GPE_CBINTERNET', GPE_CBINTERNETExit);
  BrancheOnExit('GPE_DATEEXPIRE', GPE_DATEEXPIREExit);

  SetControlVisible('GPE_CBNUMCTRL', (not CopieNumCB2Ctrl));
  SetControlVisible('TGPE_CBNUMCTRL', (not CopieNumCB2Ctrl));
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnClose ;
begin
  Inherited ;
  if Erreur then
  begin
    LastError := (-1);
    LastErrorMsg := '';
    Erreur := False;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILCB.OnCancel () ;
begin
  Inherited ;
end ;

{==============================================================================================}
{================================== TOF_MFOCBFROMKB ===========================================}
{==============================================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnUpdate ;
var
  Stg, sNumCB, sLib, sDate, sNom, sPrenom: string;
  Ctrl: TControl;
begin
  Inherited ;
  sNumCB := Trim(GetControlText('GPE_CBINTERNET'));
  if sNumCB = '' then
  begin
    if not SansClient then
    begin
      // pour demander la validation par l'utilisateur
      LastError := 1;
      Erreur := True;
    end;
    Stg := GetControlText('CONTENUCB');
    //Stg := '%B4970100000314563^AZERTY/JEANNE.MRS     ^04101234567890';4970100000314563=04109019000000000000?

    if Copy(Stg, 1, 2) = '%B' then
    begin
      System.Delete(Stg, 1, 2);
      sNumCB := ReadTokenPipe(Stg, '^');
      sLib := ReadTokenPipe(Stg, '^');
      sDate := Copy(Stg, 3, 2) + '/20' + Copy(Stg, 1, 2);
      sNom := ReadTokenPipe(sLib, '/');
      sPrenom := ReadTokenPipe(sLib, '.');
    end else
    begin
      ReadTokenST(Stg);
      sNumCB := ReadTokenPipe(Stg, '=');
      sDate := Copy(Stg, 3, 2) + '/20' + Copy(Stg, 1, 2);
      sLib := '';
      sNom := '';
      sPrenom := '';
    end;

    if sNumCB <> '' then
    begin
      SetControlText('GPE_CBINTERNET', sNumCB);
      SetControlText('GPE_TYPECARTE', CodeCardType(sNumCB));
      SetControlText('GPE_DATEEXPIRE', sDate);
      SetControlText('GPE_CBLIBELLE', Trim(sLib) + ' ' + sNom + ' ' + sPrenom);

      Ctrl := GetControl('GPE_TIERS');
      if Assigned (Ctrl) and (Ctrl is THCritMaskEdit) and (sLib <> '') and (not SansClient) then
      begin
        Stg := 'T_LIBELLE=' + Trim(sNom);
        if sPrenom <> '' then Stg := Stg + ';T_PRENOM=' + Trim(sPrenom);

        if ctxFO in V_PGI.PGIContexte then
        begin
          Stg := AGLLanceFiche('MFO', 'FOCLIMUL_MODE', Stg, '', 'SELECTION') ;
          SetControlText('GPE_TIERS', Stg);
        end else
          DispatchRecherche(THCritMaskEdit(Ctrl), 2, Stg, '', '');
      end;
    end;
  end;

  if (sNumCB <> '') and (not Erreur) then
  begin
    if (Assigned(Ecran)) and (Ecran is TFVierge) then
    begin
      TFVierge(Ecran).Retour := sNumCB +';'+ GetControlText('GPE_TYPECARTE')
         +';'+ GetControlText('GPE_DATEEXPIRE') +';'+ GetControlText('GPE_CBLIBELLE')
         +';'+ GetControlText('GPE_TIERS') +';';
    end;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnArgument (S : String ) ;
var
  Stg, NomChamp, ValeurChamp: string;
  Ind: integer;
  Ctrl: TControl;
begin
  Inherited ;
  Erreur := False;
  SansClient := False;
  Stg := S;
  repeat
    Stg := Trim(ReadTokenSt(S));
    if UpperCase(Stg) = 'SANSCLIENT' then
    begin
      SansClient := True;
      SetControlVisible('GPE_TIERS', False);
      SetControlVisible('TGPE_TIERS', False);
    end else
    if Stg <> '' then
    begin
      Ind := Pos('=', Stg);
      if Ind <> 0 then
      begin
        NomChamp := UpperCase(Copy(Stg, 1, (Ind - 1)));
        ValeurChamp := Trim(Copy(Stg, (Ind + 1), MaxInt));
        if NomChamp = 'CAPTION' then
        begin
          Ecran.Caption := TraduireMemoire(ValeurChamp);
          UpdateCaption(Ecran);
        end else
          SetControlText(NomChamp, ValeurChamp);
      end;
    end;
  until S = '';


  if V_PGI.SAV then
  begin
   Ctrl := GetControl('CONTENUCB');
   if (Ctrl <> nil) and (Ctrl is THEdit) then
   begin
     THEdit(Ctrl).PasswordChar := #0;
     THEdit(Ctrl).Font := Ecran.Font;
   end;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnClose ;
begin
  Inherited ;
  if Erreur then
  begin
    LastError := 1;
    Erreur := False;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/11/2003
Modifié le ... : 12/11/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOCBFROMKB.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MFODETAILCB ] ) ;
  registerclasses ( [ TOF_MFOCBFROMKB ] ) ;
end.

