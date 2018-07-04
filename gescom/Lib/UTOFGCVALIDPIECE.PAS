{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 29/03/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCVALIDPIECE ()
Mots clefs ... : TOF;GCVALIDPIECE
*****************************************************************}
Unit UTOFGCVALIDPIECE ;

Interface

Uses UTOF;

Type
  TOF_GCVALIDPIECE = Class (TOF)
  end ;

Implementation

Uses Classes, M3FP, Vierge, TiersUtil, Forms;

procedure AGLVALIDPIECE_AfficheRisque( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  AfficheRisqueClient(totof.LaTob);
end;

Initialization
  registerclasses ( [ TOF_GCVALIDPIECE ] ) ;
  RegisterAglProc( 'VALIDPIECE_AfficheRisque', TRUE , 1, AGLVALIDPIECE_AfficheRisque);
end.
