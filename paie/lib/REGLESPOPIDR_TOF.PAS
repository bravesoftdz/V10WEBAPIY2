{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : REGLESPOPIDR_TOF ()
Mots clefs ... : TOF;REGLESPOPIDR_TOF
*****************************************************************
pt1    07/05/07 : NA  Contrôle validité du type de population IDR}

Unit REGLESPOPIDR_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     pgoutils2,
     pgpopuloutils,
     UTOF ;

Type
  TOF_REGLESPOPIDR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    pred : string;
    procedure OnClickpopactive(Sender: TObject);
  end ;

Implementation

procedure TOF_REGLESPOPIDR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REGLESPOPIDR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_REGLESPOPIDR.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/05/2007
Modifié le ... :   /  /    
Description .. : Chargement des données
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLESPOPIDR.OnLoad ;
var
populok : boolean;

begin
  Inherited ;
// pt1

   populok := canusepopulation('IDR');
   if not populok then
   begin
    PgiInfo('Le type de population "IDR" (prédéfini '+pred+') n''est pas valide.', Ecran.caption);
    setcontrolenabled('BINSERT', false);
   end
   else
   setcontrolenabled('BINSERT', true);

   if  TCheckBox(GetControl('POPACTIVE'))<>nil then
      if (GetControlText('POPACTIVE')='X') then setcontroltext('XX_WHERE', 'PPC_PREDEFINI = "'+pred+'"')
      else setcontroltext('XX_WHERE', '');
   
// fin pt1
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLESPOPIDR.OnArgument (S : String ) ;
var
st: string;
popactive : TCheckBox;
begin
  Inherited ;
  // initialise le plus dans la tablette PGPOPULATIONIDR

  // recherche du prédéfini
  pred := GetPredefiniPopulation('IDR');
  popactive:=TCheckBox(GetControl('POPACTIVE'));
  if popactive.checked = true then st := ' and PPC_PREDEFINI = "'+pred+'"' else st := '';
  setcontrolproperty('PPR_POPULATION', 'plus', st);

  popactive:=TCheckBox(GetControl('POPACTIVE'));
  popactive.OnClick:=OnClickpopactive;
end ;

procedure TOF_REGLESPOPIDR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_REGLESPOPIDR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_REGLESPOPIDR.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_REGLESPOPIDR.Onclickpopactive(sender: Tobject) ;
var
popactive : Tcheckbox;
st : string;
begin
 popactive:=TCheckBox(GetControl('POPACTIVE'));
 if popactive.checked = true then st := ' and PPC_PREDEFINI = "'+pred+'"' else st := '';
 setcontrolproperty('PPR_POPULATION', 'plus', st);
 TFMul(Ecran).BCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_REGLESPOPIDR ] ) ;
end.
