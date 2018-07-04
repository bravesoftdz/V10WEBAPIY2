unit UTofAfActiviteMulBM;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,
      UTob,Grids,EntGC,
{$IFDEF EAGLCLIENT}
     eMul,
{$ELSE}
      db,dbTables,HDB,  Mul,
{$ENDIF}
      DicoAF,Saisutil,Hstatus,M3FP,UTofAfBaseCodeAffaire,Afactivite,
      activiteutil,
      AfUtilArticle,TraducAffaire;
Type
     TOF_AfActiviteMulBM = Class (TOF_AFBASECODEAFFAIRE)
     public
        procedure OnLoad ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose ; override ;
        procedure Onupdate; override ;
        procedure TypeAppelChargeCle(var AvActionFiche:TActionFiche);override;
     private
        titre : string;

    END ;

     const
	// libellés des messages de la TOF  AfActiviteMul
	TexteMsgAffaire: array[1..2] of string 	= (
          {1}        'Vous devez renseigner une nouvelle nature'
          {2}        ,'Code Prestation invalide'
                     );


implementation

{ TOF_AfActiviteMulBM}


procedure TOF_AfActiviteMulBM.OnLoad;
Begin
 Inherited;

End;
procedure TOF_AfActiviteMulbm.OnClose;

Begin
 Inherited;
End;

procedure TOF_AfActiviteMulbm.TypeAppelChargeCle(var AvActionFiche:TActionFiche);
begin
AvActionFiche := taConsult;
end;

procedure TOF_AfActiviteMulBM.OnArgument(stArgument : String );
Var
  Critere, Champ, valeur  : String;
  x : integer;
  zaff,zdateD,zdateF,zrep : string;
  ComboTypeArticle : THMultiValComboBox;
begin
zdated := '';
zdatef := '';
titre := Ecran.caption;
// Recup des critères
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'AFA_AFFAIRE' then Zaff := Valeur;
        if Champ = 'AFA_REPRISEACTIV'then Zrep := Valeur;
        if Champ = 'ACT_DATEACTIVITE'then ZDateD := Valeur;
        if Champ = 'ACT_DATEACTIVITE_'then ZDateF := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
SetControlText('ACT_AFFAIRE',zaff);

// Init du code affaire dans la tof ancêtre
Inherited;

// PL le 27/06/01 : si c'est pour le taConsult, j'ai créé une fonction TypeAppelChargeCle qui résoud le problème
//ChargeCleAffaire(Nil,THEDIT(GetControl('ACT_AFFAIRE1')), THEDIT(GetControl('ACT_AFFAIRE2')),
//THEDIT(GetControl('ACT_AFFAIRE3')),THEdit(GetControl('ACT_AVENANT')), Nil, taConsult,zaff,False);

if ((zdated = '') and (getcontroltext('ACT_DATEACTIVITE') <> '')) then
   zdated :=  getcontroltext('ACT_DATEACTIVITE');

   if ((zdatef = '') and(getcontroltext('ACT_DATEACTIVITE_') <> '')) then
   zdatef :=  getcontroltext('ACT_DATEACTIVITE_');

if (ZDateD = '') then
  ZDateD := DateToStr(V_PGI.DateEntree);

if (ZDateF = '') then
  ZDateF := DateToStr(V_PGI.DateEntree);


SetControlText('ACT_DATEACTIVITE',zdateD);
SetControlText('ACT_DATEACTIVITE_',zdateF);
//SetControlText('ACT_ACTIVITEREPRIS','F');   // PL le 09/09/03 : le champ n'existe pas sur la fiche
SetControlText('ACT_TYPEACTIVITE','BON');
  //mcd 05/03/02
ComboTypeArticle:=THMultiValComboBox(GetControl('ACT_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;
if ComboTypeArticle.Text='' then ComboTypeArticle.Text:=PlusTypeArticleText;
Tfmul(Ecran).FiltreDisabled:=true;  //mcd 23/05/03

end;

procedure TOF_AfActiviteMulBM.Onupdate;
begin
Inherited;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end;



/////////////// Procedure appellé par le bouton Validation //////////////




Initialization
registerclasses([TOF_AfActiviteMulBM]);
end.


