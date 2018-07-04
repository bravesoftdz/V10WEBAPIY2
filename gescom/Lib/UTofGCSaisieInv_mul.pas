{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/08/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCSAISIEINV_MUL ()
Mots clefs ... : TOF;GCSAISIEINV_MUL
*****************************************************************}
Unit UTofGCSaisieInv_mul ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
     eMul,
{$ELSE}
     Mul,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF, M3FP ;

Type
  TOF_GCSAISIEINV_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure SupprimeListeInv(CodeListe:string) ;
  end ;

Implementation

const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Etes-vous sûr de vouloir supprimer cet inventaire ?'
                     );

procedure TOF_GCSAISIEINV_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCSAISIEINV_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCSAISIEINV_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCSAISIEINV_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCSAISIEINV_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  if (ctxMode in V_PGI.PGIContexte) then SetControlVisible('BDelete', True);
end ;

procedure TOF_GCSAISIEINV_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_GCSAISIEINV_MUL.SupprimeListeInv(CodeListe:string) ;
var SQL : string ;
begin
   if (PGIAsk(TraduireMemoire(TexteMessage[1]),Ecran.Caption) = mrNo) then exit;

   SQL:='delete from LISTEINVENT where GIE_CODELISTE="'+CodeListe+'"';
   ExecuteSQL(SQL) ;

   SQL:='delete from LISTEINVLIG where GIL_CODELISTE="'+CodeListe+'"';
   ExecuteSQL(SQL) ;

   SQL:='delete from LISTEINVLOT where GLI_CODELISTE="'+CodeListe+'"';
   ExecuteSQL(SQL) ;
end;

procedure AGLSupprimeListeInv(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                else exit;
if (TOTOF is TOF_GCSAISIEINV_MUL) then TOF_GCSAISIEINV_MUL(TOTOF).SupprimeListeInv(Parms[1])
                                  else exit;
end;

procedure InitTOFSAISIEINV_MUL ;
begin
  RegisterAglProc('SupprimeListeInv', True , 1, AGLSupprimeListeInv) ;
End ;

Initialization
  registerclasses ( [ TOF_GCSAISIEINV_MUL ] ) ;
  InitTOFSAISIEINV_MUL ;
end.
