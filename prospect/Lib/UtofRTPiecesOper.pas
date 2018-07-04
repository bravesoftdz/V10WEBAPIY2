{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTMULPIECESOPER ()
Mots clefs ... : TOF;RTPiecesOper
*****************************************************************}
Unit UtofRTPiecesOper ;

Interface

Uses Classes,
     forms,HCtrls, 
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      Fe_Main,Mul,
{$ENDIF}

{$IFDEF VER150}
Variants,
{$ENDIF}

     UTOF,ParamSoc,UtilSelection,UtilOperation,HTB97 ;

Type
  TOF_RTPiecesOper = Class (TOF)
    Private
      stAction : string;
      procedure BPiece_OnClick(Sender: TObject);
      procedure BDelete_OnClick(Sender: TObject);
    Public
      procedure OnArgument (S : String ) ; override ;
  end ;

Function RTLanceFiche_RTPiecesOper(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

Function RTLanceFiche_RTPiecesOper(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTPiecesOper.OnArgument (S : String ) ;
var F:TForm;
    Critere,ChampMul,ValMul : string;
    x : integer;
begin
  Inherited ;
  if GetParamSocSecur('SO_RTGESTINFOS00D',True) = True then
  begin
    F:=TForm (Ecran);
    MulCreerPagesCL(F,'NOMFIC=PIECES');
  end;
  if Assigned(GetControl('BPIECE')) then
     TToolbarButton97(GetControl('BPIECE')).OnClick := BPIECE_OnClick;
  if Assigned(GetControl('BDELETE')) then
     TToolbarButton97(GetControl('BDELETE')).OnClick := BDELETE_OnClick;
  repeat
      Critere:=(ReadTokenSt(S));
      if Critere <> '' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='ACTION' then StAction := ValMul;
           end;
        if Critere = 'NOSUPPRLIEN' then SetControlVisible ('BDELETE',False);
        end;
  until Critere='';
  if stAction = 'CONSULTATION' then
  begin
    SetControlVisible ('BDELETE',False);
    SetControlVisible ('BPIECE',False);
  end;
end ;

procedure TOF_RTPiecesOper.BPiece_OnClick(Sender: TObject);
begin
  RTAttachePieceOper(GetControlText('GP_OPERATION'));
  TFMul(Ecran).BChercheClick(Nil);
end;

procedure TOF_RTPiecesOper.BDelete_OnClick(Sender: TObject);
begin
  if GetField('GP_NATUREPIECEG') <> Null then
  begin
    RTSuppressionLienPieceOper(GetControlText('GP_OPERATION'),GetField('GP_NATUREPIECEG'),
         GetField('GP_SOUCHE'),GetField('GP_NUMERO'),GetField('GP_INDICEG'));
    TFMul(Ecran).BChercheClick(Nil);
  end;
end;

Initialization
  registerclasses ( [ TOF_RTPiecesOper ] ) ;
end.

