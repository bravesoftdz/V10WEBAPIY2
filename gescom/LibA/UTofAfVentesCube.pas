unit UTofAfventescube;


interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,paramsoc,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,
{$ENDIF}
      HCtrls,UTOF,AFTableauBord,FactUtil,
      AfUtilArticle,Dicobtp,HEnt1,EntGC,HTB97,utofbaseetats,
      UTofAfBaseCodeAffaire, UTOB, Stat,M3FP,UTofAfBasePiece_Mul;

Type

  //TOF_Afventescube = Class (TOF_AFBASECODEAFFAIRE) // mcd 10/12/02 pour restriction sur nature piece idem autres fiches
   TOF_Afventescube = Class (TOF_AFBASEPIECE_MUL)
   			procedure OnArgument (stArgument : string); override;
        procedure OnLoad; override;
        procedure OnClose; override;
        private
        procedure NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);override ;
     END;
 Procedure AFLanceFiche_VentesCUbe(Range,Argument:string);

implementation



//******************************************************************************
//**************** TOB d'affichage du tableau de bord **************************
//******************************************************************************
procedure TOF_Afventescube.OnArgument(stArgument : String );
Var //CC   : TCheckBox;
    ComboTypeArticle : THMultiValComboBox;
    Critere, Champ, valeur,zori  : String;
    x : integer;
Begin
Inherited;
zori := 'P';
Critere:=(Trim(ReadTokenSt(stArgument)));   // TABLE:GP;NATURE:FPR;NOCHANGE_NATUREPIECE
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end
        else
        	champ := critere;
        if Champ = 'AFFAIRE' then Zori:='A';
        if Champ = 'PIECE' 	 then Zori:='P';
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
if (zori = 'A') then
    Begin      // fait sur lignes associées affaire et propo uniquement + dates affaire
    Ecran.Caption := TraduitGA('Statistiques sur lignes d''affaires');
    SetControlText('GL_NATUREPIECEG',GetParamSoc('SO_AFNATAFFAIRE'));
    ThMultiValCOmboBox(GetControl('GL_NATUREPIECEG')).Datatype:='GCNATUREPIECEG'; //mcd 04/07/03
    ThMultiValCOmboBox(GetControl('GL_NATUREPIECEG')).Plus:=' AND (GPP_NATUREPIECEG="'+
      GetParamSoc('SO_AFNATAFFAIRE')+'" OR GPP_NATUREPIECEG="'+ GetParamSoc('SO_AFNATPROPOSITION')+'")'; //mcd 04/07/03
    SetControlText('TGL_DATEPIECE','Facturation Affaire');
    SetControlVIsible ('GL_DATEPIECE',false); SetControlEnabled('GL_DATEPIECE',false);
    SetControlVIsible ('GL_DATEPIECE_',false);SetControlEnabled('GL_DATEPIECE_',false);
    SetControlVIsible ('AFF_DATEDEBGENER',true);SetControlEnabled('AFF_DATEDEBGENER',true);
    SetControlVIsible ('AFF_DATEFINGENER',true);SetControlEnabled('AFF_DATEFINGENER',true);
    SetControlText('GL_DATEPIECE',DateToStr(idate1900));
    SetControlText('GL_DATEPIECE_',DateToStr(idate2099));
    End
  else
    Begin   // fait sur nature de pièce vente + date pièce
    Ecran.Caption := TraduitGA('Statistiques de vente');
	  SetControlText('GL_NATUREPIECEG','FAC;AVC');
    SetControlText('TGL_DATEPIECE','Dates Pièces');
    SetControlVIsible ('GL_DATEPIECE',true); SetControlEnabled('GL_DATEPIECE',true);
    SetControlVIsible ('GL_DATEPIECE_',true);SetControlEnabled('GL_DATEPIECE_',true);
    SetControlVIsible ('AFF_DATEDEBGENER',false);SetControlEnabled('AFF_DATEDEBGENER',false);
    SetControlVIsible ('AFF_DATEFINGENER',false);SetControlEnabled('AFF_DATEFINGENER',false);
    SetControlText('AFF_DATEDEBGENER',DateToStr(idate1900));
    SetControlText('AFF_DATEFINGENER',DateToStr(idate2099));
    End;




  //mcd 05/03/02
ComboTypeArticle:=THMultiValComboBox(GetControl('GL_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;
if ComboTypeArticle.Text='' then ComboTypeArticle.Text:=PlusTypeArticleText;

UpdateCaption(Ecran);
End;


// Récupération du paramétrage et lancement du tob viewer
procedure TOF_Afventescube.OnLoad;
Var CodeAffaire, CodeTiers{, ListeChampsSel} : string;
    //TOBTB : TOB;
    //CC   : TCheckBox;
    //Part0, Part1  ,Part2 ,Part3,Part4 : string;
BEGIN
CodeAffaire := ''; CodeTiers := '';

END;

procedure TOF_Afventescube.NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('GL_TIERS'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers_:=THEdit(GetControl('GL_TIERS_'));
end;

procedure TOF_Afventescube.OnClose;
BEGIN

END;

 Procedure AFLanceFiche_VentesCUbe(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFVENTES_CUBE',Range,'',Argument);
end;

Initialization
registerclasses([TOF_Afventescube]);
end.
