{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGREGROUP_MUL ()
Mots clefs ... : TOF;PGREGROUP_MUL
***************************************************************** }

Unit UTofMULPGREGROUP ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     HDB,
     mul,
     fe_main,
{$else}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     Hqry,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF, htb97 ; 

Type
  TOF_MULPGREGROUP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    typeregroupement : string;
     procedure GrilleDblClick(sender: Tobject);
     procedure Binsertclick(sender: Tobject);
  end ;

Implementation

procedure TOF_MULPGREGROUP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULPGREGROUP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULPGREGROUP.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Chargement
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPGREGROUP.OnLoad ;
var
st : string;
begin
  st := '';
  if typeregroupement = 'MAB' then
  st := 'AND PRB_TYPEREGROUP = "'+typeregroupement+'"'
  else
  st := 'AND PRB_TYPEREGROUP <> "MAB"';

  setcontroltext('XX_WHERE', st);

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPGREGROUP.OnArgument (S : String ) ;
var
  st : string;
  Binsert : TToolBarButton97;
  {$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
begin
  Inherited ;
    typeregroupement := readtokenst(S);
    if typeregroupement = 'MAB' then
    begin
    setcontrolvisible('PRB_TYPEREGROUP', false);
    setcontrolvisible('TPRB_TYPEREGROUP', false);
    end;
   st := 'AND CO_ABREGE LIKE "%'+typeregroupement+'%"';
   SetControlProperty('PRB_THEMEREGROUP','Plus',St);

   {$IFNDEF EAGLCLIENT}
   Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
   Liste := THGrid(GetControl('FListe'));
{$ENDIF}
   if Liste <> nil then
   Liste.OnDblClick := GrilleDblClick;

   Binsert := TToolBarButton97(GetControl('BINSERT'));
   if Binsert <> nil then Binsert.OnClick := BinsertClick;
end ;



procedure TOF_MULPGREGROUP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MULPGREGROUP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULPGREGROUP.OnCancel () ;
begin
  Inherited ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Doubleclick
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPGREGROUP.GrilleDblClick(sender: Tobject);

begin
  AGLLancefiche('PAY', 'PGREGROUP', '',Getfield('PRB_PREDEFINI') +';'+ getfield('PRB_NODOSSIER') +';'+ Getfield('PRB_CODEREGROUP'),'ACTION=MODIFICATI0N' + ';'+ typeregroupement);
  TFMul(Ecran).BCherche.Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Création
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPGREGROUP.BinsertClick(sender: Tobject);
begin

  AGLLanceFiche('PAY', 'PGREGROUP', '','', 'ACTION=CREATION' + ';' + typeregroupement);
  TFMul(Ecran).BCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_MULPGREGROUP ] ) ;
end.
