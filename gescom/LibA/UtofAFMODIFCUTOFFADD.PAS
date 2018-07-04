{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 07/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFCUTOFFADD ()
Mots clefs ... : TOF;AFMODIFCUTOFFADD
*****************************************************************}
Unit UtofAFMODIFCUTOFFADD ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
Maineagl,
{$ELSE}
       db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
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
     dicobtp, utilCutoff ;

Type
  TOF_AFMODIFCUTOFFADD = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    public
    sTypeEclat:string;
    bSortie:boolean;
    GModeEclat: T_ModeEclat;
  end ;

function AFLanceFiche_ModifCutOffAdd(Argument:String):string;

Implementation

procedure TOF_AFMODIFCUTOFFADD.OnArgument (S : String ) ;
var Critere, Champ, valeur: String;
    X : integer;
    sPanelCourant:string;
begin
  Inherited ;
TFVierge(Ecran).Retour := '';
bSortie:=true;
sTypeEclat:='';

Critere:=(Trim(ReadTokenSt(S)));
  While (Critere <>'') do
  begin
    X:=pos(':',Critere);
    if X<>0 then
       begin
       Champ:=copy(Critere,1,X-1);
       Valeur:=Copy (Critere,X+1,length(Critere)-X);
       end;
    if (Champ = 'TYPE')   then
        begin
        if (Valeur = 'FAC')   then
            begin
            sTypeEclat:='FAC';
            // Mode d'eclatement
            GModeEclat:=DetermineModeEclatFact;
            end;
        if (Valeur = 'CVE') then
            begin
            sTypeEclat:='CVE';
            // Mode d'eclatement
            GModeEclat:=DetermineModeEclatCutOff;
            end;
        end;
    Critere:=(Trim(ReadTokenSt(S)));
  end;

if sTypeEclat='FAC' then
    begin
    sPanelCourant:= 'PANELFAC';
    SetControlVisible('PANELFAC', true);
    SetControlVisible('PANELDEV', false);
    SetControlProperty('PANELFAC', 'left', THEdit(GetControl('PANELDEV')).left);
    SetControlProperty('PANELFAC', 'top', THEdit(GetControl('PANELDEV')).top);
    end
else
    begin
    sPanelCourant:= 'PANELDEV';
    SetControlVisible('PANELFAC', false);
    SetControlVisible('PANELDEV', true);
    end;

case GModeEclat of
        tmeRessArt :begin
                SetControlVisible('ASSISTANT', true);
                SetControlVisible('LBLASSISTANT', true);
                SetControlVisible('TYPEARTICLE', false);
                SetControlVisible('LBLTYPEARTICLE', false);
                SetControlVisible('ARTICLE', true);
                SetControlVisible('LBLARTICLE', true);
                end;
        tmeArticle :begin
                SetControlVisible('ASSISTANT', false);
                SetControlVisible('LBLASSISTANT', false);
                SetControlVisible('TYPEARTICLE', false);
                SetControlVisible('LBLTYPEARTICLE', false);
                SetControlVisible('ARTICLE', true);
                SetControlVisible('LBLARTICLE', true);
                SetControlProperty(sPanelCourant, 'left', THEdit(GetControl('ARTICLE')).left);
                SetControlProperty('ARTICLE', 'left', THEdit(GetControl('ASSISTANT')).left);
                SetControlProperty('LBLARTICLE', 'left', THEdit(GetControl('LBLASSISTANT')).left);
                TForm(Ecran).width := TForm(Ecran).width - THEdit(GetControl('ASSISTANT')).width;
                end;
        tmeRessource :begin
                SetControlVisible('ASSISTANT', true);
                SetControlVisible('LBLASSISTANT', true);
                SetControlVisible('TYPEARTICLE', false);
                SetControlVisible('LBLTYPEARTICLE', false);
                SetControlVisible('ARTICLE', false);
                SetControlVisible('LBLARTICLE', false);
                SetControlProperty(sPanelCourant, 'left', THEdit(GetControl('ARTICLE')).left);
                TForm(Ecran).width := TForm(Ecran).width - THEdit(GetControl('ARTICLE')).width;
                end;
        tmeRessTypA :begin
                SetControlVisible('ASSISTANT', true);
                SetControlVisible('LBLASSISTANT', true);
                SetControlVisible('TYPEARTICLE', true);
                SetControlVisible('LBLTYPEARTICLE', true);
                SetControlVisible('ARTICLE', false);
                SetControlVisible('LBLARTICLE', false);
                end;
        tmeTypeArticle :begin
                SetControlVisible('ASSISTANT', false);
                SetControlVisible('LBLASSISTANT', false);
                SetControlVisible('TYPEARTICLE', true);
                SetControlVisible('LBLTYPEARTICLE', true);
                SetControlVisible('ARTICLE', false);
                SetControlVisible('LBLARTICLE', false);
                SetControlProperty(sPanelCourant, 'left', THEdit(GetControl('ARTICLE')).left);
                SetControlProperty('TYPEARTICLE', 'left', THEdit(GetControl('ASSISTANT')).left);
                SetControlProperty('LBLTYPEARTICLE', 'left', THEdit(GetControl('LBLASSISTANT')).left);
                TForm(Ecran).width := TForm(Ecran).width - THEdit(GetControl('ASSISTANT')).width;
                end;
end;

if sTypeEclat='FAC' then
    TForm(Ecran).width := TForm(Ecran).width - THEdit(GetControl('PANELDEV')).width + THEdit(GetControl('PANELFAC')).width;

TToolBarButton97(GetControl('HelpBtn')).Left := TForm(Ecran).width - TToolBarButton97(GetControl('HelpBtn')).width -20;
TToolBarButton97(GetControl('BFerme')).Left := TToolBarButton97(GetControl('HelpBtn')).Left - TToolBarButton97(GetControl('BFerme')).width -4;
TToolBarButton97(GetControl('BValider')).Left := TToolBarButton97(GetControl('BFerme')).Left - TToolBarButton97(GetControl('BValider')).width -4;

if sTypeEclat='FAC' then
    begin
    ChangeMask(THNumEdit(GetControl('PVFACT')), V_PGI.OkDecP, '') ;
    end
else
    begin
    ChangeMask(THNumEdit(GetControl('FAE')), V_PGI.OkDecP, '') ;
    ChangeMask(THNumEdit(GetControl('AAE')), V_PGI.OkDecP, '') ;
    ChangeMask(THNumEdit(GetControl('PCA')), V_PGI.OkDecP, '') ;
    end;

end ;


procedure TOF_AFMODIFCUTOFFADD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnUpdate ;
var
bIncomplet:boolean;
Critere1,Critere2, Lib1, Lib2 :string;
begin
Inherited ;
nextprevcontrol(ecran);
bIncomplet:=false;
bSortie:=true;
Critere1:='';
Critere2:='';

case GModeEclat of
        tmeRessArt :begin
                Critere1 := GetControlText('ASSISTANT');
                Critere2 := GetControlText('ARTICLE');
                Lib1 := GetControlText('LIBELLEASSISTANT');
                Lib2 := GetControlText('LIBELLEARTICLE');
                if (GetControlText('ASSISTANT')='') or (GetControlText('ARTICLE')='') then
                    bIncomplet:=true;
                end;
        tmeArticle :begin
                Critere2:=GetControlText('ARTICLE');
                Lib1 := '';
                Lib2 := GetControlText('LIBELLEARTICLE');
                if (GetControlText('ARTICLE')='') then
                    bIncomplet:=true;
                end;
        tmeRessource :begin
                Critere1:=GetControlText('ASSISTANT');
                Lib1 := GetControlText('LIBELLEASSISTANT');
                Lib2 := '';
                if (GetControlText('ASSISTANT')='') then
                    bIncomplet:=true;
                end;
        tmeRessTypA :begin
                Critere1:=GetControlText('ASSISTANT');          
                Critere2:=GetControlText('TYPEARTICLE');
                Lib1 := GetControlText('LIBELLEASSISTANT');
                Lib2 := THValComboBox(GetControl('TYPEARTICLE')).Text;
                if (GetControlText('ASSISTANT')='') or (GetControlText('TYPEARTICLE')='') then
                    bIncomplet:=true;
                end;
        tmeTypeArticle :begin
                Critere2:=GetControlText('TYPEARTICLE');
                Lib1 := '';
                Lib2 := THValComboBox(GetControl('TYPEARTICLE')).Text;
                if (GetControlText('TYPEARTICLE')='') then
                    bIncomplet:=true;
                end;
end;


if bIncomplet then
    begin
    bSortie:=false;
    PGIInfoAF('Veuillez saisir les éléments clés de la nouvelle ligne.', Ecran.Caption);
    exit;
    end;

if sTypeEclat='FAC' then
    begin
    TFVierge(Ecran).Retour := 'C1='+ Critere1 + ';C2='+ Critere2 + ';L1='+ Lib1 + ';L2='+ Lib2 + ';PVFACT='+ GetControlText('PVFACT');
    end
else
    begin
    TFVierge(Ecran).Retour := 'C1='+ Critere1 + ';C2='+ Critere2 + ';L1='+ Lib1 + ';L2='+ Lib2 + ';FAE='+ GetControlText('FAE')+';AAE='+ GetControlText('AAE')+';PCA='+GetControlText('PCA');
    end;

end ;

procedure TOF_AFMODIFCUTOFFADD.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnClose ;
begin
if not bSortie then cancel;

Inherited ;
end ;

function AFLanceFiche_ModifCutOffAdd(Argument:String):string;
begin
Result:=AGLLanceFiche ('AFF','AFMODIFCUTOFFADD','','',Argument);
end;


Initialization
  registerclasses ( [ TOF_AFMODIFCUTOFFADD ] ) ;
end.
