{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/10/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPIECETIERS_MUL ()
Mots clefs ... : TOF;BTPIECETIERS_MUL
*****************************************************************}
Unit BTPIECETIERS_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
{$ENDIF}
		 AGLInit,
     uTob,
     Menus,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTOFComm,
     UTOF ;

Type
  TOF_BTPIECETIERS_MUL = Class (tTOFComm)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    PieceBTP,PieceNEg : string;
    NewPiece : boolean;
    procedure ControleChamp(Champ, Valeur: String);
    procedure ChargePopDoc;
    procedure PieceBTPClick (Sender : TObject);
    procedure PieceNEGClick (Sender : TObject);
  end ;

Implementation
uses PARAMSOC,AGLInitRT;

procedure TOF_BTPIECETIERS_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.OnArgument (S : String ) ;
var Critere  : string;
    ValMul   : string;
    ChampMul : string;
    x        : integer;
begin
  fMulDeTraitement := True;
  Inherited ;
  NewPiece := false;
	PieceBTP := GetParamSocSecur('SO_BTCHOIXPIECEPROPOS','DBT');
  PieceNEG := GetParamSocSecur('SO_BTCHOIXPIECENEG','DE');
  SetControltext('GP_NATUREPIECEG',PieceNeg+';'+PieceBTP+';');
  ChargePopDoc;
  Repeat
    Critere:=uppercase(ReadTokenSt(S)) ;
    valMul := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
      end
      else
        ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
    end;
  until Critere='';

end ;

procedure TOF_BTPIECETIERS_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECETIERS_MUL.ControleChamp(Champ, Valeur: String);
begin
	if Champ = 'NOUVELLEPIECE' then newPiece := true;
end;

procedure TOF_BTPIECETIERS_MUL.ChargePopDoc;
var PP : TPopupMenu;
		XX : TMenuItem;
begin
  PP := TPopupMenu(GetControl('POPDOCS'));
  if PP <> nil then
  begin
		XX := TMenuItem.Create (PP);
    XX.Name := 'MnPIECEBTP';
    XX.Caption := RechDom('GCNATUREPIECEG',PieceBTP,false);
    XX.OnClick := PieceBTPClick;
    PP.Items.Add(XX);
    //
		XX := TMenuItem.Create (PP);
    XX.Name := 'MnPIECENEG';
    XX.Caption := RechDom('GCNATUREPIECEG',PieceNEG,false);
    XX.OnClick := PieceNEGClick;
    PP.Items.Add(XX);
  end;
end;

procedure TOF_BTPIECETIERS_MUL.PieceBTPClick(Sender: TObject);
begin
  if NewPiece then
  begin
    if CreePieceRT([GetControlText('GP_TIERS'),GetControlText('PERSPECTIVE'),PieceBTP],3) <> '' then Close else RefreshDB();
  end else
  begin
    AglLanceFiche ('RT','RTPIECETIERS_MUL','GP_TIERS='+GetControltext('GP_TIERS')+';GP_NATUREPIECEG='+PieceBTP+
              ';GP_PERSPECTIVE=0;PERSPECTIVE='+GetControltext('GP_PERSPECTIVE'),'','NOUVELLEPIECE');
    RefreshDB();
  end;
end;

procedure TOF_BTPIECETIERS_MUL.PieceNEGClick(Sender: TObject);
begin
  if NewPiece then
  begin
    if CreePieceRT([GetControlText('GP_TIERS'),GetControlText('PERSPECTIVE'),PieceNEG],3) <> '' then Close else RefreshDB();
  end else
  begin
    AglLanceFiche ('RT','RTPIECETIERS_MUL','GP_TIERS='+GetControltext('GP_TIERS')+';GP_NATUREPIECEG='+PieceNEG+
              ';GP_PERSPECTIVE=0;PERSPECTIVE='+GetControltext('GP_PERSPECTIVE'),'','NOUVELLEPIECE');
    RefreshDB();
  end;

end;

Initialization
  registerclasses ( [ TOF_BTPIECETIERS_MUL ] ) ; 
end.
