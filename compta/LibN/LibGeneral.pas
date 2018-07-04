unit LibGeneral;

interface

uses classes, uTOB, SysUtils;

procedure RemplirNatureSens ( Gene,Auxi : string ; var Nature,Sens : string ) ;
function  GeneNatureSisco2VersPGI ( NatureSisco : string ) : string;

type  TGeneral = class ( TOB )
        private
          procedure SetLibelle ( St : string );
        public
          constructor   Create ( T : TOB ); reintroduce ; overload ;
          destructor    Destroy ; override ;
          procedure     Charge ( Compte : string );
          procedure     InitNouveau ( Compte : string );
          procedure     SetNature ( Nature : string );
        published
          property Libelle : string write SetLibelle;
      end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  Hent1,
  Ent1;

{ TGeneral }

procedure TGeneral.Charge(Compte: string);
begin
  SelectDB ('"'+Compte+'"',nil);
end;

constructor TGeneral.Create ( T : TOB ) ;
begin
  inherited Create ('GENERAUX', T, -1) ;
end;

destructor TGeneral.Destroy;
begin
  inherited Destroy ;
end;

procedure TGeneral.InitNouveau( Compte: string );
var Libelle, Nature, Sens : string;
begin
  Compte := BourreLaDonc(Compte,fbGene);
  if Length(Compte)<=3 then exit ;
  PutValue ('G_GENERAL',Compte);
  RemplirNatureSens(Compte,'',Nature,Sens) ;
  if ((Nature='') or (Sens='')) then Exit ;
  PutValue('G_NATUREGENE',Nature) ;
  if ((Nature='COC') or (Nature='COF') or (Nature='COS') or (Nature='COD')) then
    PutValue('G_COLLECTIF','X') ;
  PutValue('G_SENS',Sens) ;
  // Par défaut, on renseigne le libellé. Utiliser la propriété Libelle pour le renseigner
  // avec une valeur bien définie
  Libelle := TraduireMemoire ('CREE PAR CEGID');
  PutValue('G_ABREGE',Copy(Libelle,1,17)) ;
  PutValue ('G_LIBELLE', Libelle );
  PutValue ('G_CONFIDENTIEL','0');
  PutValue ('G_VENTILABLE1','-');
  PutValue ('G_VENTILABLE2','-');
  PutValue ('G_VENTILABLE3','-');
  PutValue ('G_VENTILABLE4','-');
  PutValue ('G_VENTILABLE5','-');
end;

Procedure RemplirNatureSens ( Gene,Auxi : String ; Var Nature,Sens : String ) ;
Var CC,C2,C3 : Char ;
BEGIN
CC:=Gene[1] ; C2:=Gene[2] ; C3:=Gene[3] ;
Nature:='' ; Sens:='' ;
Case CC of
   '1' : BEGIN Nature:='DIV' ; Sens:='C' ; END ;
   '2' : BEGIN
         Nature:='IMO' ; Sens:='D' ;
         if C2 in ['8','9'] then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '3' : BEGIN
         Nature:='DIV' ; Sens:='D' ;
         if C2='9' then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '4' : BEGIN
         Nature:='DIV' ; if Auxi<>'' then Nature:='COD' ;
         Case C2 of
            '0' : BEGIN
                  Sens:='C' ; if C3='9' then Sens:='D' ;
                  if Auxi<>'' then Nature:='COF' ;
                  END ;
            '1' : BEGIN
                  Sens:='D' ; if C3='9' then Sens:='C' ;
                  if Auxi<>'' then Nature:='COC' ;
                  END ;
            '8' : BEGIN Sens:='D' ; END ;
             else BEGIN Sens:='C' ; END ;
            END ;
         END ;
   '5' : BEGIN
         Nature:='DIV' ; Sens:='D' ;
         if C2='1' then BEGIN Nature:='BQE' ; Sens:='D' ; END else
         if C2='3' then BEGIN Nature:='CAI' ; Sens:='D' ; END else
         if C2='9' then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '6' : BEGIN Nature:='CHA' ; Sens:='D' ; END ;
   '7' : BEGIN Nature:='PRO' ; Sens:='C' ; END ;
   END ;
END ;

function  GeneNatureSisco2VersPGI ( NatureSisco : string ) : string;
begin
  if NatureSisco = 'G' then Result := ''  // La nature est déjà correctement renseignée par RemplirNatureSens
  else if NatureSisco = 'F' then Result := 'COF'
  else if NatureSisco = 'C' then Result := 'COC';
end;

procedure TGeneral.SetLibelle(St: string);
begin
  St := Trim (St);
  if St <> '' then
  begin
    PutValue('G_LIBELLE',Copy ( St, 1, 35 ) );
    PutValue('G_ABREGE',Copy ( St, 1, 17 ) );
  end else
  begin
    PutValue('G_LIBELLE',TraduireMemoire('CREE PAR CEGID'));
    PutValue('G_ABREGE',TraduireMemoire('CREE PAR CEGID'));
  end;
end;

procedure TGeneral.SetNature(Nature: string);
begin
  if Nature = '' then exit;
  PutValue ('G_NATUREGENE', Nature);
  if Copy (Nature,1 ,2 ) = 'CO' then PutValue ('G_COLLECTIF','X')
  else PutValue ('G_COLLECTIF','-');
end;

end.
