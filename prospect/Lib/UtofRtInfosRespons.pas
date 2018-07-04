{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTINFOSRESPONS ()
Mots clefs ... : TOF;RTINFOSRESPONS
*****************************************************************}
Unit UtofRTInfosRespons ;

Interface

Uses
     Classes,
 {$ifdef GIGI}
     entdp, Paramsoc,
 {$ENDIF GIGI}
     UTOF,Vierge,Windows,htb97 ;

Type
  TOF_RTINFOSRESPONS = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
       procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

Implementation

procedure TOF_RTINFOSRESPONS.OnArgument (S : String ) ;
begin
  Inherited ;
  if (ecran <> Nil)then TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
{$IFDEF GIGI} //mcd 06/05/2005 pour être cohérent avec menu
 if  (Vh_DP.SeriaMessagerie) then
   begin  //si agenda sérialisé dans le bureau, on ne le voit pas en GI
   SetControlVisible ('BPLANNING',False)
   end;
 if not GetParamsocSecur('SO_AFRTCHAINAGE',False) then
   begin
   SetControlVisible ('BRAPPELCH',False)
   end;
{$ENDIF}
end ;

procedure TOF_RTINFOSRESPONS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState); 
begin
  Inherited ;
  Case Key of
    VK_RETURN :
        if TToolBarButton97(GetControl('bFerme')) <> nil then
           TToolBarButton97(GetControl('bFerme')).OnClick(Sender);
  end;
end ;
procedure TOF_RTINFOSRESPONS.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_RTINFOSRESPONS ] ) ; 
end.
