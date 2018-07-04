unit UTofGcArticle_Rech;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,
      HCtrls,HMsgBox,UTOF,M3FP, ent1,utilGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fe_Main,
{$ENDIF}
      HEnt1 ;

type
     TOF_GCARTICLE_RECH = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
     END ;

implementation

procedure TOF_GCARTICLE_RECH.OnArgument(Arguments : String) ;
var DroitCreat : boolean ;
    Nbr : integer ;
Begin
inherited ;
Nbr := 0;
if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else Nbr := Nbr + 1;
{$IFNDEF CCS3}
if (Nbr = 0) then
{$ENDIF}
   SetControlVisible('PZONESLIBRES',False) ;


DroitCreat:=ExJaiLeDroitConcept(TConcept(gcArtCreat),False) ;
if (Not DroitCreat) then
  begin
  // Suppression des boutons d'insertion et de duplication
  SetControlVisible('BINSERT', False) ;
  end ;
End ;

Initialization
registerclasses([TOF_GCARTICLE_RECH]);

end.
 