{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCVENTES ()
Mots clefs ... : TOF;GCVENTES
*****************************************************************}
Unit UtofAchatCube ;

Interface

Uses StdCtrls, Controls, Classes, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UtilGC ;

Type
  TOF_GCACHAT_CUBE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_GCACHAT_CUBE.OnArgument (S : String ) ;
Var i : integer ;
    Nam,Titre : String ;
begin
  Inherited ;
for i:=1 to 5 do BEGIN Nam:='GL_LIBREART'+IntToStr(i); GCTitreZoneLibre(Nam,Titre); SetControlCaption('T'+Nam,Titre); END;
//   for i:=1 to 3 do BEGIN Nam:='GP_LIBREFOU'+IntToStr(i) ; GCTitreZoneLibre(Nam,Titre) ; SetControlCaption('T'+Nam,Titre); end;
if (ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) 
 then begin
 THMultiValComboBox(GetControl('GL_NATUREPIECEG')).Plus :=  AfPlusNatureAchat;
 SetControlText('GL_NATUREPIECEG','FF');
 end;

end ;

Initialization
  registerclasses ( [ TOF_GCACHAT_CUBE ] ) ;
end.

