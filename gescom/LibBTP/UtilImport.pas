unit UtilImport;

interface

uses Hctrls;

procedure AffecteTablette(var ChaineInit : string;Valeur:string ; AvecNomDebut:boolean=true;AvecParenthese :boolean=true);
function ChampTablette (Chaine:string) : string;

implementation

procedure AffecteTablette(var ChaineInit : string;Valeur:string ; AvecNomDebut:boolean=true;AvecParenthese :boolean=true);
var Debut,longueur,nbdec,typeD,Champ,Tablette : string;
    NomTablette,STrav :string;
begin
if AvecNomDebut then
   begin
   NomTablette := Copy ( ChaineInit,1,Pos('=',ChaineInit)-1);
   Strav := copy (ChaineInit,Pos('=',ChaineInit)+2,255);
   Strav := copy (Strav,1,pos(')',strav)-1);
   end else strav := ChaineInit;
Debut:=READTOKENST (strav);
Longueur:=READTOKENST (strav);
Nbdec:=READTOKENST (strav);
TypeD:=READTOKENST (strav);
Champ:=READTOKENST (strav);
if AvecNomDebut then ChaineInit := NomTablette+'=' else ChaineInit := '';
if AvecParenthese then ChaineInit := ChaineInit+'(';
CHaineInit := Chaineinit+Debut+';'+Longueur+';'+Nbdec+';'+TypeD+';'+Champ+';'+Valeur;
if AvecParenthese then ChaineInit := ChaineInit+')';
end;

function ChampTablette (Chaine:string) : string;
var tmpch : string;
    Debut,longueur,nbdec,typeD,Champ,Tablette : string;
begin
if pos (')',chaine) > 0  then Tmpch := copy (chaine,1,pos(')',chaine)-1)
                         else TmpCh := Chaine;
Debut:=READTOKENST (TmpCh);
Longueur:=READTOKENST (TmpCh);
Nbdec:=READTOKENST (TmpCh);
TypeD:=READTOKENST (TmpCh);
Champ:=READTOKENST (TmpCh);
result := READTOKENST (TmpCh);
end;


end.
 