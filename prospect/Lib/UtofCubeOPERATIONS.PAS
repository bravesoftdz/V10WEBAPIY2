{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OPERATIONS_CUBE ()
Mots clefs ... : TOF;
*****************************************************************}
Unit UtofCubeOPERATIONS ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
{$else}
     MainEAGL,
{$ENDIF}
     forms,
     UTOF,
     UtilSelection,ParamSoc,Cube ;

Function RTLanceFiche_CubeOPERATIONS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_CubeOPERATIONS = Class (TOF)
     public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

Function RTLanceFiche_CubeOPERATIONS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_CubeOPERATIONS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CubeOPERATIONS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CubeOPERATIONS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CubeOPERATIONS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CubeOPERATIONS.OnArgument (S : String ) ;
var F : TForm;
begin
  Inherited ;
  if S <> 'GRF' then
    begin
    F := TForm (Ecran);
    if GetParamSocSecur('SO_RTGESTINFOS002',False) = True then
        MulCreerPagesCL(F,'NOMFIC=RTOPERATIONS');
    if GetParamSocSecur ('SO_RTLIENOPERATIONPIECE',False) then
      begin
      SetControlVisible('PPIECE',True) ;
//      SetControlText ('GP_VENTEACHAT','VEN');
      TFCUBE(Ecran).FromSQL :='OPERATIONS LEFT JOIN RTINFOS002 ON ROP_OPERATION = RD2_CLEDATA' +
                          ' LEFT JOIN PIECE ON ROP_OPERATION = GP_OPERATION AND GP_VENTEACHAT = "VEN"'+
                          ' LEFT JOIN RTINFOS00D ON RDD_CLEDATA=CAST(GP_NBTRANSMIS AS varchar(7))';
      if GetParamSocSecur('SO_RTGESTINFOS00D',True) = True then
        begin
        F:=TForm (Ecran);
        MulCreerPagesCL(F,'NOMFIC=PIECES');
        end;
      end;
    end;
  SetControlText ('XX_WHERE','ROP_OPERATION<>"MODELES D''ACTIONS"');
{$Ifdef GIGI}
 if (GetControl('GP_REPRESENTANT') <> nil) then  SetControlVisible('GP_REPRESENTANT',false);
 if (GetControl('TGP_REPRESENTANT') <> nil) then  SetControlVisible('TGP_REPRESENTANT',false);
 if (GetControl('YTC_REPRESENTANT2') <> nil) then  SetControlVisible('YTC_REPRESENTANT2',false);
 if (GetControl('YTC_REPRESENTANT2_') <> nil) then  SetControlVisible('YTC_REPRESENTANT2_',false);
 if (GetControl('YTC_REPRESENTANT3') <> nil) then  SetControlVisible('YTC_REPRESENTANT3',false);
 if (GetControl('YTC_REPRESENTANT3_') <> nil) then  SetControlVisible('YTC_REPRESENTANT3_',false);
{$endif GIGI}
end ;

procedure TOF_CubeOPERATIONS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CubeOPERATIONS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CubeOPERATIONS.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CubeOPERATIONS ] ) ;
end.
