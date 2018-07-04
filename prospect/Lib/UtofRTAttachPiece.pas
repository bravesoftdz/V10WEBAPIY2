{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTATTACHPIECE ()
Mots clefs ... : TOF;RTATTACHPIECE
*****************************************************************}
Unit UtofRTAttachPiece ;

Interface

Uses Classes,
     forms, 
     UTOF,ParamSoc,UtilSelection ;

Type
  TOF_RTATTACHPIECE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_RTATTACHPIECE.OnArgument (S : String ) ;
var F:TForm;
begin
  Inherited ;
  if GetParamSocSecur('SO_RTGESTINFOS00D',True) = True then
  begin
    F:=TForm (Ecran);
    MulCreerPagesCL(F,'NOMFIC=PIECES');
  end;
  if Pos('ORIGINEINTERVENTION',S) <> 0 then
    SetControlProperty('GP_NATUREPIECEG','plus','AND GPP_PIECESAV="X"');
end ;

Initialization
  registerclasses ( [ TOF_RTATTACHPIECE ] ) ;
end.

