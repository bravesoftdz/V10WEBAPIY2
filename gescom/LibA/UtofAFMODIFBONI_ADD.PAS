{***********UNITE*************************************************
Auteur  ...... :MCD
Créé le ...... : 25/08/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFBONI_ADD ()
Mots clefs ... : TOF;AFMODIFBONI_ADD
*****************************************************************}
Unit UtofAFMODIFBONI_ADD ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
Maineagl,
{$ELSE}
       db,  dbTables,FE_Main,
{$ENDIF}
     forms,
     ent1,
     sysutils,
     ComCtrls,
     HCtrls,
     HTB97,
     HEnt1,
     HMsgBox,
     vierge,
     UTOF,
     dicoAF ;

Type
  TOF_AFMODIFBONI_ADD = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    public
    bMemlookup : boolean;
    sTypeEclat:string;
    bSortie:boolean;
  end ;

function AFLanceFiche_ModifBoni_Add(Argument:String):string;

Implementation

procedure TOF_AFMODIFBONI_ADD.OnArgument (S : String ) ;
begin
  Inherited ;
bMemlookup := V_PGI.LookUpLocate;

TFVierge(Ecran).Retour := '';
bSortie:=true;
sTypeEclat:='';

V_PGI.LookUpLocate := true;
SetControlProperty('ARTICLE', 'Plus', 'GROUP BY APB_CODEARTICLE ORDER BY APB_CODEARTICLE');

end ;



procedure TOF_AFMODIFBONI_ADD.OnUpdate ;
var
bIncomplet:boolean;
Critere1,Critere2 :string;
begin
Inherited ;
nextprevcontrol(ecran);
bIncomplet:=false;
bSortie:=true;
Critere1:='';
Critere2:='';
 //mcd
Critere1 := GetControlText('ASSISTANT');
Critere2 := GetControlText('ARTICLE');
if (GetControlText('ASSISTANT')='') or (GetControlText('ARTICLE')='') then
    bIncomplet:=true;
if bIncomplet then
    begin
    bSortie:=false;
    PGIInfoAF('Veuillez saisir les éléments clés de la nouvelle ligne.', Ecran.Caption);
    exit;
    end;
TFVierge(Ecran).Retour := 'C1='+ Critere1 + ';C2='+ Critere2 + ';BONI='+ GetControlText('BONI');
end ;


procedure TOF_AFMODIFBONI_ADD.OnClose ;
begin
if not bSortie then cancel;

V_PGI.LookUpLocate := bMemlookup;
Inherited ;
end ;

function AFLanceFiche_ModifBoni_Add(Argument:String):string;
begin
Result:=AGLLanceFiche ('AFF','AFMODIFBONI_ADD','','',Argument);
end;


Initialization
  registerclasses ( [ TOF_AFMODIFBONI_ADD ] ) ;
end.
