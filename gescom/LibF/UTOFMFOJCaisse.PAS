{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 31/07/2001
Description .. : Source TOF de la FICHE : JCAISSE
Suite ........ : Affichage de l'état de la caisse (journée ouverte ou fermée,
Suite ........ : communication, ...)
Mots clefs ... : TOF;UTOFMFOJCAISSE;FO
*****************************************************************}
unit UTOFMFOJCaisse;

interface

uses
  Classes, Graphics, sysutils,
  {$IFNDEF EAGLCLIENT}
  dbTables,
  {$ENDIF}
  HCtrls, HEnt1, UTOF, UTOB, ParamSoc;

type
  TOF_MFOJCAISSE = class(TOF)
    procedure OnLoad; override;
    procedure OnArgument(stArgument: string); override;
  end;

implementation
uses
  EntGC,
  {$IFDEF TOXCLIENT}
  UTox, UToxNet,
  {$ENDIF}
  FOUtil;

///////////////////////////////////////////////////////////////////////////////////////
//  OnLoad :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOJCAISSE.OnLoad;
var sCaisse, sNumZ, Etat, Stg: string;
  Color: TColor;
  QQ: TQuery;
  TOBJCaisse: TOB;
  {$IFDEF TOXCLIENT}
  WhenStart: TDateTime;
  {$ENDIF}
begin
  inherited;
  // Lecture de la table JOURSCAISSE pour la caisse et la journée courante
  TOBJCaisse := TOB.Create('JOURSCAISSE', nil, -1);
  //sCaisse := FOCaisseCourante ;
  //sNumZ := IntToStr(FOGetNumZCaisse(sCaisse)) ;
  sCaisse := GetControlText('GJC_CAISSE');
  sNumZ := GetControlText('GJC_NUMZCAISSE');
  Stg := 'select * from JOURSCAISSE where GJC_CAISSE="' + sCaisse + '" and GJC_NUMZCAISSE=' + sNumZ;
  QQ := OpenSQL(Stg, True);
  if not QQ.EOF then TOBJCaisse.SelectDB('', QQ);
  Ferme(QQ);
  TOBJCaisse.PutEcran(Ecran);
  // Ajout des champs complémentaires
  Etat := TOBJCaisse.GetValue('GJC_ETAT');
  if Etat = 'OUV' then
  begin
    Color := clGreen;
    Stg := 'Ouverte';
  end else
    if Etat = 'CPR' then
  begin
    Color := $00006FDD;
    Stg := 'Clôturée provisoirement';
  end else
    if Etat = 'CDE' then
  begin
    Color := clRed;
    Stg := 'Clôturée définitivement';
  end else
  begin
    Color := clBlack;
    Stg := 'Non ouverte';
  end;
  FoSetFontColor(Ecran, 'LGJC_ETAT', Color);
  SetControlText('LGJC_ETAT', TraduireMemoire(Stg));
  // Format des dates d'ouverture et de fermeture
  if TOBJCaisse.GetValue('GJC_DATEOUV') = IDate1900 then SetControlText('GJC_DATEOUV', '');
  if TOBJCaisse.GetValue('GJC_DATEFERME') = IDate1900 then SetControlText('GJC_DATEFERME', '');
  // Vendeur qui a lancé l'ouverture et la fermeture
  if (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISIE') = '-') and (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISLIG') = '-') then
  begin
    SetControlVisible('GJC_VENDOUV', False);
    SetControlVisible('GJC_VENDFERME', False);
  end;
  TOBJCaisse.Free;
  {$IFDEF TOXCLIENT}
  //  Etat de la liaison FO - BO
  if AglStatusTox(WhenStart) then
  begin
    Stg := TraduireMemoire('Les échanges sont activés') + ' '
      + TraduireMemoire('depuis le') + ' ' + DateTimeToStr(WhenStart);
    Color := clGreen;
  end else
  begin
    Stg := TraduireMemoire('Les échanges ne sont pas activés');
    Color := clRed;
  end;
  FoSetFontColor(Ecran, 'LETATTOX', Color);
  SetControlText('LETATTOX', Stg);
  // Etat du ToxServeur
  if AglToxCommuniCam(UST_RUNNING, nil) then
  begin
    Stg := TraduireMemoire('Le serveur de communication est démarré');
    Color := clGreen;
  end else
  begin
    Stg := TraduireMemoire('Le serveur de communication est arrêté');
    Color := clRed;
  end;
  FoSetFontColor(Ecran, 'LETATTOXSRV', Color);
  SetControlText('LETATTOXSRV', Stg);
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnArgument :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOJCAISSE.OnArgument(stArgument: string);
var FoBo: Boolean;
  Ind: Integer;
  sCritere, sNomChamp, sValeur: string;
begin
  inherited;
  {$IFDEF TOXCLIENT}
  FoBo := True;
  {$ELSE}
  FoBo := False;
  {$ENDIF}
  repeat
    sCritere := Trim(ReadTokenSt(stArgument));
    if sCritere = 'SANSTOX' then FoBo := False;
    if sCritere <> '' then
    begin
      Ind := Pos('=', sCritere);
      if Ind <> 0 then
      begin
        sNomChamp := Copy(sCritere, 1, Ind - 1);
        sValeur := Copy(sCritere, Ind + 1, Length(sCritere));
        if sNomChamp <> '' then SetControlText(sNomChamp, sValeur);
      end;
    end;
  until sCritere = '';
  SetControlVisible('BValider', False);
  SetControlVisible('BVERSION', False);
  if FoBo then
  begin
    sValeur := Trim(GetParamSoc('SO_GCFOCAISREFTOX'));
    FoBo := ((sValeur <> '') and (sValeur = GetControlText('GJC_CAISSE')));
  end;
  SetControlVisible('BFOBO', FoBo);
end;

initialization
  registerclasses([TOF_MFOJCAISSE]);
end.
