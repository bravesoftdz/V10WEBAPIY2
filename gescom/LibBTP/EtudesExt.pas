unit EtudesExt;

interface

uses SysUtils,UTOB,EntGc,BTPUtil;
procedure determineMiniMaxi (schaine : string;var mini : integer;var maxi: integer);

type TEtudesExt = class
     private
     TOBImp : TOB;
     function RecupTob(Atob:TOB) : integer;
     function ChercheEmplacPlage (TOBSTrDoc : TOB): integer;
     public
     procedure RempliStruct (var TOBStrDoc : TOB);
     procedure createStruct (var TOBStrDoc: TOB; TOBEtude: TOB; RowEtude: integer);
end;


implementation


procedure determineMiniMaxi (schaine : string;var mini : integer;var maxi: integer);
begin
mini := -1;
Maxi := -1;
if pos ('-',schaine) <= 0 then exit;
mini := strtoint (copy(schaine,1,pos('-',schaine)-1));
maxi := strtoint (copy(schaine,pos('-',schaine)+1,255));
end;

{ TEtudeExt }

procedure TEtudesExt.createStruct(var TOBStrDoc: TOB; TOBEtude: TOB; RowEtude: integer);
var TOBL : TOB;
begin
if TOBEtude <> nil then
   begin
   TOBSTRDOC.putvalue('BSD_TIERS',TOBEtude.GetValue('AFF_TIERS'));
   TOBSTRDOC.putvalue('BSD_AFFAIRE',TOBEtude.GetValue('AFF_AFFAIRE'));
   TOBSTRDOC.putvalue('BSD_ORDRE',TOBEtude.detail[RowEtude-1].GetValue('BDE_ORDRE'));
   TOBSTRDOC.putvalue('BSD_TYPE',TOBEtude.detail[RowEtude-1].GetValue('BDE_TYPE'));
   TOBSTRDOC.putvalue('BSD_NATUREDOC',TOBEtude.detail[RowEtude-1].GetValue('BDE_NATUREDOC'));
   TOBSTRDOC.putvalue('BSD_INDICE',TOBEtude.detail[RowEtude-1].GetValue('BDE_INDICE'));
   end;
// Plage de prise en compte
{if (TOBEtude<> nil) and (TOBEtude.detail[RowEtude-1].getValue('BDE_QUALIFNAT') <> 'PRINC1') then
   begin
   TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
   TOBL.addchampsupValeur('NATURE','000');
   TOBL.addchampsupValeur('FEUILLE','');
   TOBL.addchampsupValeur('INDICE','0');
   TOBL.addchampsupValeur('VALEUR','');
   TOBL.addchampsupValeur('OBLIG','X');
   end;}
// colonne N° Prix
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','001');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');
// colonne N° Designation
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','002');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');
// colonne N° Quantite
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','003');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');
// colonne N° Unite
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','004');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');
// colonne N° P.U
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','005');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');
// colonne N° Montant
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','006');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');
// colonne N° Code Article/Ouvrage
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','007');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','');
{
// Nom de la feuille
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,-1);
TOBL.addchampsupValeur('NATURE','007');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE','0');
TOBL.addchampsupValeur('VALEUR','');
}
end;

function TEtudesExt.RecupTob(Atob: TOB):integer;
begin
result := 0;
TOBImp := ATob;
end;

function TEtudesExt.ChercheEmplacPlage (TOBSTrDoc : TOB): integer;
var indice : integer;
begin
result := -1;
for Indice := 0 to TOBStrDoc.detail.count -1 do
    begin
    if TOBSTRDOC.detail[Indice].getValue('NATURE') > '000' then BEGIN result := Indice;break; END;
    end;
end;

procedure TEtudesExt.RempliStruct(var TOBStrDoc: TOB);
var Indice : integer;
    TOBL,TOBI : TOB;
begin
if TOBStrDoc.GetValue('BSD_DESCRIPT') = '' then exit;
TOBLoadFromBuffer (TOBSTRDoc.GetValue('BSD_DESCRIPT'),RecupTob);
if TOBIMP.detail.count > 0 then
for Indice := 0 to TOBIMP.detail.count -1 do
    begin
    TOBI := TOBImp.detail[Indice];
    if (not TOBI.FieldExists ('FEUILLE')) then TOBI.addchampsupValeur ('FEUILLE','');
    if (not TOBI.FieldExists ('INDICE')) then TOBI.addchampsupValeur ('INDICE','0');
    if (not TOBI.FieldExists ('OBLIG')) then TOBI.addchampsupValeur ('OBLIG','X');
    TOBL := TobStrDoc.findfirst (['NATURE','INDICE'],[TOBI.GetValue('NATURE'),TOBI.GetValue('INDICE')],true);
    if (TOBL = nil) and (TOBI.GetValue('NATURE')='000') then
       begin

       TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,ChercheEmplacPlage (TOBSTrDOC));
       TOBL.addchampsupValeur('NATURE','000');
       TOBL.addchampsupValeur('FEUILLE',TOBI.GetValue('FEUILLE'));
       TOBL.addchampsupValeur('INDICE',strtoint(TOBI.GetValue('INDICE')));
       TOBL.addchampsupValeur('VALEUR',TOBI.GetValue('VALEUR'));
       TOBL.addchampsupValeur('OBLIG',TOBI.GetValue('OBLIG'));
       end else
       begin
       if TOBL <> nil then
          begin
          TOBL.PutValue('NATURE',TOBI.GetValue('NATURE'));
          TOBL.putvalue('FEUILLE',TOBI.GetValue('FEUILLE'));
          TOBL.putvalue('INDICE',strtoint(TOBI.GetValue('INDICE')));
          TOBL.PutValue('VALEUR',TOBI.GetValue('VALEUR'));
          TOBL.PUTVAlue('OBLIG',TOBI.GetValue('OBLIG'));
          end;
       end;
    end;
TOBImp.free;
end;

end.
