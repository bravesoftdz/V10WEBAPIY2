{***********UNITE*************************************************
Auteur  ...... : CHARRAIX
Créé le ...... : 06/06/2001
Modifié le ... : 06/06/2001
Description .. : Source TOF de la TABLE : RESHISTOVALO ()
Mots clefs ... : TOF;RESHISTOVALO
*****************************************************************}
Unit UTofResHistoValo ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOM,Dialogs,TiersUtil,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}                     
      FE_main,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      EntGC, Dicobtp,HTB97,Vierge, UTOB, UTOF, AGLInit,Ent1,Paramsoc,AffaireUtil;

Type
  TOF_RESHISTOVALO = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    //LaRessource  : String;
    Modification : boolean;
    Validation   : boolean;
    BValider     : TToolbarButton97;
    LSTChamps    : String;
    T            : TOB;
    Parms        : String;
    procedure StockParam;
    procedure BvaliderClick(Sender : Tobject);
    Procedure VerifModifications;
  end ;
Function AFLanceFiche_RevalActivParam(Argument:string):variant;

Implementation



procedure TOF_RESHISTOVALO.OnArgument (S : String ) ;
var i : integer;
BEGIN
  Inherited ;
  LSTChamps := 'ARS_DATEPRIX,ARS_TAUXREVIENTUN,ARS_PVHT,ARS_PVTTC,ARS_PVHTCALCUL';
  Modification  := False;
  Validation := False;
  //
  BValider := TToolbarButton97(GetControl('BVALIDER'));
  BValider.OnClick := BValiderClick;

  //Décimales prix unitaires
  for i := 2 to 4 do
  begin
    ChangeMask(THNumEdit(GetControl('ARS_TAUXREVIENTUN'+IntToStr(i))), V_PGI.OkDecP, '');
    ChangeMask(THNumEdit(GetControl('ARS_PVHTCALCUL'+IntToStr(i))), V_PGI.OkDecP, '');
    ChangeMask(THNumEdit(GetControl('ARS_PVHT'+IntToStr(i))), V_PGI.OkDecP, '');
    ChangeMask(THNumEdit(GetControl('ARS_PVTTC'+IntToStr(i))), V_PGI.OkDecP, '');
  end;
{$IFDEF AFFAIRE}
  if (GetParamSoc ('SO_AFCLIENT') = cInClientAlgoe) then  //Couts unitaires
  begin
    SetControlVisible ('TARS_PVTTC1',true);
    SetControlVisible ('ARS_TAUXUNIT2',true);
    SetControlVisible ('ARS_TAUXUNIT3',true);
    SetControlVisible ('ARS_TAUXUNIT4',true);
  end;
{$ENDIF}
END;


procedure TOF_RESHISTOVALO.OnClose ;
BEGIN
  VerifModifications;
  T.free;
  If Modification then
    begin
      If (Validation = False) and (PgiAskAf('Des données ont été modifiées, voulez-vous conserver ces modifications ?',TitreHalley) = MrNo ) then exit;
      If (StrToDate(ThEdit(GetControl('ARS_DATEPRIX2')).text) > StrToDate(ThEdit(GetControl('ARS_DATEPRIX3')).text))
          and (StrToDate(ThEdit(GetControl('ARS_DATEPRIX3')).text) > StrToDate(ThEdit(GetControl('ARS_DATEPRIX4')).text)) then
        StockParam
      else
       begin
        PgiInfo('Toutes les dates doivent être différentes et doivent respecter un ordre chronologique',TitreHalley);
        exit;
       end;
    end
  else parms := '1';
  //On fait le retour dans le script de la fiche
  TFVierge(Ecran).retour := parms;
  Thetob := Latob;
  Inherited ;
END ;

procedure TOF_RESHISTOVALO.StockParam;
BEGIN
  if modification = True then
     begin
       parms := '0';
       LaTob.GetEcran(Ecran);
     end
  else parms := '1';
  Inherited ;
END ;

procedure TOF_RESHISTOVALO.BvaliderClick;
begin
  Validation := True;
end;

procedure TOF_RESHISTOVALO.VerifModifications;
var
{parms,}cel,champs,ListeChamps : string;
i : integer;
BEGIN
 ListeChamps := LSTChamps;
 While ListeChamps<>'' do
  begin
   cel := ReadTokenPipe(ListeChamps,',');
   champs := cel;
   For i:=2 to 4 do
    begin
      if cel = 'ARS_DATEPRIX' then
       begin
        If StrToDate(ThEdit(GetControl(cel+IntToStr(i))).text) <> VarToDateTime(LaTob.GetValue(champs+IntToStr(i))) then Modification := True;
       end
      else
      If GetControlText(cel+IntToStr(i)) <> LaTob.GetValue(champs+IntToStr(i)) then Modification := True;
    end;
  end;
END;

Function AFLanceFiche_RevalActivParam(Argument:string):variant;
begin
result:=AGLLanceFiche('AFF','RESHISTOVALO','','',Argument);
end;

Initialization
  registerclasses ( [ TOF_RESHISTOVALO ] ) ; 
end.


 
