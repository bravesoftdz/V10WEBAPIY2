{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/11/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PRESRECALCUL ()
                 Saisie de la date pour le recalcul des compteurs de présence
Mots clefs ... :
***************************************************************** }

Unit UTOF_PGRecalculPresence ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     sysutils,
     HStatus,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Vierge,
     HTB97,
     ed_tools,
     Pgpresence,
     UTOF ;

Type
  TOF_PGRECALCULPRESENCE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     daterecalcul : TDatetime;
     procedure recalculcompteur(Sender : Tobject);
  end ;

Implementation

procedure TOF_PGRECALCULPRESENCE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGRECALCULPRESENCE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGRECALCULPRESENCE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGRECALCULPRESENCE.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 05/11/2007
Modifié le ... :   /  /
Description .. : ON ARGUMENT
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGRECALCULPRESENCE.OnArgument (S : String ) ;
var
Btn     : TToolBarButton97;
libellemodif, libellemodif2,datevalidite : string;
lib, lib2 : THlabel;

begin
  Inherited ;
  libellemodif := readtokenst(S);
  libellemodif2 := readtokenst(S);
  datevalidite := readtokenst(S);
  daterecalcul := strtodate(datevalidite);
  Lib := THLabel(GetControl('LIBELLEMODIF'));
  SetControlText ('DATERECALCUL', datevalidite);
  Lib.Caption := libellemodif;
  Lib2 := THLabel(GetControl('LIBELLEMODIF2'));
  Lib2.Caption := libellemodif2;

   Btn := TToolBarButton97(GetControl('B_VALIDER'));
  if btn<>nil then btn.OnClick := Recalculcompteur;

end ;

procedure TOF_PGRECALCULPRESENCE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGRECALCULPRESENCE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGRECALCULPRESENCE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 05/11/2007
Modifié le ... :   /  /    
Description .. : Mise à jour de la date de recalcul des compteurs de présence
Mots clefs ... :
*****************************************************************}
procedure TOF_PGRECALCULPRESENCE.recalculcompteur(sender : Tobject);
var
datesaisie : Tdatetime;
Begin

datesaisie := StrToDate(getcontroltext('DATERECALCUL'));;
if (datesaisie < daterecalcul) then
begin
PGIBOX('La date du recalcul doit être >= date de validité', Ecran.caption);
setfocuscontrol('DATERECALCUL');
exit;
end;



CompteursARecalculer(Datesaisie);
Ecran.Close;
end;


Initialization
  registerclasses ( [ TOF_PGRECALCULPRESENCE ] ) ;
end.
