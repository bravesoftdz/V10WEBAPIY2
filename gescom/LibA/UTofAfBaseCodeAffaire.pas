{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : début 2000
Modifié le ... :   /  /
Description .. : Source TOF  : AFBASECODEAFFAIRE
Ancêtre permettant la gestion automatique
Mots clefs ... : TOF;AFTRADUCCHAMPLIBRE
*****************************************************************}
unit UTofAfBaseCodeAffaire;

interface
uses  StdCtrls,
	  Controls,
      Classes,
{$IFDEF EAGLCLIENT}
      Utob,
      eMul,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      mul,
{$ENDIF}

      forms,
      sysutils,
      ComCtrls,
      ParamSoc,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      entgc,
      Ent1,

{$IFDEF AFFAIRE}
      AffaireUtil,
{$ENDIF}

{$IFDEF BTP}
      CalcOleGenericBTP, AppelsUtil,
{$ENDIF}
      factutil,
      DicoBTP,
      SaisUtil,
      HTB97,
      UTOFAFTRADUCCHAMPLIBRE;



Type
     TOF_AFBASECODEAFFAIRE = Class (TOF_AFTRADUCCHAMPLIBRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); virtual;
        procedure TypeAppelChargeCle(var AvActionFiche:TActionFiche);virtual;
        procedure bSelectAff1Click(Sender: TObject);virtual;
        procedure bSelectAff2Click(Sender: TObject);
        procedure SelectionAffaire(EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4 :THEdit);
        procedure bEffaceAff1Click(Sender: TObject);virtual;
        procedure bSelectConClick(Sender: TObject);
        procedure bEffaceAff2Click(Sender: TObject);
        procedure EffaceAffaire(EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4 :THEdit);
        procedure bSelectAffMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
        procedure GereChampConfidentialite(TConfLoc : ThEdit; deux : string);
     Private
      BCreatAff : Boolean;
      Status : string;
      fEtatAffaire : string;
    procedure BRechResponsable(Sender: TObject);
     public
      bEffaceAff1    :  TToolBarButton97;
      bEffaceAff2    :  TToolBarButton97;
      bSelectAff1    :  TToolBarButton97;
      bSelectAff2    :  TToolBarButton97;
      bSelectCon     :  TToolBarButton97;
      EditTiers      :  THEdit;
      EditTiers_     :  THEdit;
      EditAff        :  THEdit;
      EditAff0       :  THEdit;
      EditAff1       :  THEdit;
      EditAff2       :  THEdit;
      EditAff3       :  THEdit;
      EditAff4       :  THEdit;
      EditAff_       :  THEdit;
      EditAff0_      :  THEdit;
      EditAff1_      :  THEdit;
      EditAff2_      :  THEdit;
      EditAff3_      :  THEdit;
      EditAff4_      :  THEdit;
      Tconf,Tconf_   :  THEdit;
      bChangeTiers   :  Boolean;
      afActionFiche  :  TActionFiche;
      bSaisieAffaire :  boolean;
      bchangeStatut  :  boolean;
      EtatAffaire : THMultiValComboBox;
      property TheStatus : String read Status write Status;
      property TheEtatAffaire : String read fEtatAffaire write fEtatAffaire;
END ;

implementation
uses UtilRessource;
//uses GalOutil;

procedure TOF_AFBASECODEAFFAIRE.OnArgument(stArgument : String );
Var StAff  : String;
    StAff_ : String;
    Argument: string;
    critere : string;
    X : integer;
    Champ   : string;
    Valeur  : string;
    CC      : THValComboBox;
    CCE : Thedit;
Begin
Inherited;
//
// Gestion de la selection des affaires de debut et de fin de selection
//
	Argument := stArgument;
	bchangeStatut := false;
  //Status := 'AFF';
	Critere:=(Trim(ReadTokenSt(Argument)));

  While (Critere <>'') do
    BEGIN
    X:=pos(':',Critere);
    if x = 0 then X:=pos('=',Critere);
    if x<>0 then
       begin
       Champ:=copy(Critere,1,X-1);
       Valeur:=Copy (Critere,X+1,length(Critere)-X);
       end
    else
       Champ := Critere;
    if champ = 'STATUT'             then Status := valeur
    else if Champ = 'CHANGESTATUT'  then bChangeStatut:=True
    else if Champ = 'COTRAITANCE'   then
    begin
      if assigned(GetControl('XX_WHERE')) then Setcontroltext('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-")')
    end;

    Critere:=(Trim(ReadTokenSt(Argument)));
  END;
	//

  bChangeTiers := True;
  BCreatAff :=True;
  afActionFiche := taCreat;
  bSaisieAffaire:=false;

  If Ecran.name='AFPARAMPROTOAFF' then
  begin //mcd 25/09/03 pour ne pas permettre dans ce cas le changement code tiers et pas bouton efface
  BChangeTiers :=False;
  BCreatAff :=False;
  end;

  //
  //NCX 13/11/2001 Confidentialité affaires : utile dans le onargument seulement pour les QRS1 qui ne passent dans le onload (là où toutes
  // les initialisations sont faites normalement) qu'au moment du lancement de l'état... En laissant l'appel aux fonctions ici, l'utilisateur
  // verra desuite en arrivant sur l'écran de lancement d'état, la confidentialité initialisée
  //
  GereChampConfidentialite(Tconf,'');
  GereChampConfidentialite(Tconf_,'_');

  // PL le 10/12/02 : Si on a des blocages sur les états de l'affaire et qu'on a un champ
  // MultiValComboBox sur l'état, il faut qu'il ne présente pas les états bloqués
(*  EtatAffaire := THMultiValComboBox (GetControl ('AFF_ETATAFFAIRE'));
  if (EtatAffaire <> nil) and (EtatAffaire is THMultiValComboBox) then
    begin
    if (Copy(Critere,1,3) = 'CTX')  and (EtatAffaire <> nil) then
      // Réduction de la multicombo des états de l'affaire en fonction des blocages affaire
      begin
        tobEtatsBloques := nil;
        try
          tobEtatsBloques := TOBEtatsBloquesAffaire (copy (critere, 5, 3), '-', V_PGI.groupe);
          sPlus := EtatAffaire.plus;
          if (tobEtatsBloques <> nil) then
            for i:=0 to tobEtatsBloques.Detail.count - 1 do
              begin
                sPlus := sPlus  + ' AND CC_CODE<>"' + tobEtatsBloques.Detail[i].GetValue('ABA_ETATAFFAIRE') + '"';
              end;
          SetControlProperty('AFF_ETATAFFAIRE','Plus', sPlus);
        finally
          tobEtatsBloques.Free;
        end;
      end;
    end;
*)

// Ajout PA le 24/09/2001 pour passer d'autres arguments au chargeCleAffaire si l'on veut un code en création d'affaire (utilisé par acceptation + duplication)
if Pos('AFFAIREENCREATION',stArgument) > 0 then bSaisieAffaire:=True;
NomsChampsAffaire(EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, EditAff_, EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_, EditTiers, EditTiers_);
bSelectAff1 := TToolBarButton97(GetControl('BSELECTAFF1'));
bSelectAff2 := TToolBarButton97(GetControl('BSELECTAFF2'));
bSelectCon  := TToolBarButton97(GetControl('BSELECTCON'));

  if (bSelectAff1<>nil) then
    begin
    bSelectAff1.OnClick:=bSelectAff1Click;
    bSelectAff1.OnMouseDown:=bSelectAffMouseDown;
    end;

  if (bSelectCon<>nil) then
    begin
    bSelectCon.OnClick:=bSelectConClick;
    bSelectCon.OnMouseDown:=bSelectAffMouseDown;
    end;

  if (bSelectAff2<>nil) then
    begin
    bSelectAff2.OnClick:=bSelectAff2Click;
    bSelectAff2.OnMouseDown:=bSelectAffMouseDown;
    end;


  if (EditAff1<>nil) or (EditAff1_<>nil) then TypeAppelChargeCle(afActionFiche);

  StAff  :='' ;
  StAff_ :='' ;

  if Status <> '' then
     Begin
     if Editaff0  <> Nil then Editaff0.text := Status;
     if Editaff0_ <> Nil then Editaff0_.text:= Status;
     end;

  if (EditAff0 <> nil) then
     BEGIN
     if EditAff <> Nil then StAff := EditAff.Text;
   ChargeCleAffaire(EditAff0,EditAff1,EditAff2,EditAff3,EditAff4,Nil,afActionFiche,StAff,bSaisieAffaire);
   END
  else
	 ChargeCleAffaire(EditAff0,EditAff1,EditAff2,EditAff3,EditAff4,Nil,afActionFiche,StAff,bSaisieAffaire);
//
// Gestion d'un  deuxième code dans une même fiche
//
	//Editaff0_.text := Status;

	if (EditAff0_ <> Nil) then
     BEGIN
   	 if EditAff_ <> Nil then StAff_ := EditAff_.Text;
	   ChargeCleAffaire(EditAff0_,EditAff1_,EditAff2_,EditAff3_,EditAff4_,Nil,afActionFiche,StAff_,bSaisieAffaire);
  	 end
	else
  	 ChargeCleAffaire(EditAff0_,EditAff1_,EditAff2_,EditAff3_,EditAff4_,Nil,afActionFiche,StAff_,bSaisieAffaire);
//
// Gestion de l'effacement des champs de l'affaire et du tiers
//
  bEffaceAff1 := TToolBarButton97(GetControl('BEFFACEAFF1'));
if (bEffaceAff1<>nil) then bEffaceAff1.OnClick:=bEffaceAff1Click;

  bEffaceAff2 := TToolBarButton97(GetControl('BEFFACEAFF2'));
if (bEffaceAff2<>nil) then bEffaceAff2.OnClick:=bEffaceAff2Click;

  // Cacher les boutons si pas en mode affaire
  if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
   BEGIN
   if (bSelectAff1<>nil) then SetControlVisible('BSELECTAFF1',False);
   if (bSelectAff2<>nil) then SetControlVisible('BSELECTAFF2',False);
   if (bEffaceAff1<>nil) then SetControlVisible('BEFFACEAFF1',False);
   if (bEffaceAff2<>nil) then SetControlVisible('BEFFACEAFF2',False);
   END;

   // mcd 10/07/2002
  If GetControl('AFF_GENERAUTO') <> nil then
  begin
	if (ecran.name <> 'BTIMPDGD_MUL') then
  begin
  {$IFDEF BTP}
    SetControlProperty('AFF_GENERAUTO','Plus','BTP');
  {$ELSE}
  SetControlProperty('AFF_GENERAUTO','Plus','GA');
  if ctxScot in V_PGI.PGIContexte then  SetControlProperty('AFF_GENERAUTO','Plus','GA" AND CO_CODE<>"CON');
  {$ENDIF}
  end;
  end;


  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

  CC:=THValComboBox(GetControl('AFF_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

	CCE:=ThEdit(GetControl('AFF_RESPONSABLE')) ;
	if CCE<>Nil then
  begin
  	PositionneResponsableUser(CCE) ;
    CCE.OnElipsisClick := BRechResponsable;
	end;

End;



procedure TOF_AFBASECODEAFFAIRE.BRechResponsable(Sender: TObject);
Var QQ  : TQuery;
    SS  : THCritMaskEdit;
begin

  if GetParamSocSecur('SO_AFRECHRESAV', True) then
  begin
    SS := THCritMaskEdit.Create(nil);
    GetRessourceRecherche(SS,'ARS_RESSOURCE=' + ThEdit(Sender).text + ';TYPERESSOURCE=SAL', '', '');
    if (SS.Text <> THEdit(Sender).text) then
    begin
      if SS.text = '' then ss.text := ThEdit(Sender).text;
    end;
    ThEdit(Sender).text  := SS.Text;
    SS.Free;
  end
  else
    GetRessourceRecherche(ThEdit(Sender),'ARS_TYPERESSOURCE="SAL"', '', '');

end;

procedure TOF_AFBASECODEAFFAIRE.OnUpdate;
Begin
Inherited;
  //if Status = 'APP' then
  //begin
  {$IFDEF EAGLCLIENT}
  //	if (Ecran is TFMul) then ModifColAppelGrid ( TFMul(Ecran).FListe,TFMul(Ecran).Q.tq);
  {$ELSE}
  //	if (Ecran is TFMul) then ModifColAppelGrid ( TFMul(Ecran).FListe);
  {$ENDIF}
  //end
  //else
  //begin
  {$IFDEF AFFAIRE}
  {$IFDEF EAGLCLIENT}
  if (Ecran is TFMul) then ModifColAffaireGrid ( TFMul(Ecran).FListe,TFMul(Ecran).Q.tq);
  {$ELSE}
  if (Ecran is TFMul) then ModifColAffaireGrid ( TFMul(Ecran));
  {$ENDIF}
  //if (Ecran is TFMul) then TFMul(Ecran).FListe.Repaint; // PL 06/01/03 pour traduction OK
  {$ENDIF}
  //end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CHARRAIX
Créé le ...... : 09/01/2002
Modifié le ... :   /  /
Description .. : Confidentialité Mission par groupe de confidentialité, par
Suite ........ : responsable ou par assitant libre.
Mots clefs ... : CONFIDENTIALITE
*****************************************************************}
procedure  TOF_AFBASECODEAFFAIRE.GereChampConfidentialite(TConfLoc : ThEdit; deux : string);
var
i:integer;
Valinit:string;
TGroupeConf : ThMultiValComboBox;
LTGroupeConf : ThLabel;
vargroupeConf, critere, plus : String;
begin
TconfLoc:=nil;
Valinit:='';
 // mcd 12/09/01 le terme groupe confidentialité deveint groupe de travial. fait par prog
If (ThLabel(GetControl('TAFF_GROUPECONF')) <>Nil) then SetControlCaption ('TAFF_GROUPECONF','Groupe travail');
If (ThLabel(GetControl('TATBXGROUPECONF')) <>Nil) then SetControlCaption ('TATBXGROUPECONF','Groupe travail');
//gestion de l'affichage de la zone de confidentialité
TGroupeConf := ThMultiValComboBox(GetControl('AFF_GROUPECONF'));
if (TGroupeConf=nil) then
   TGroupeConf := ThMultiValComboBox(GetControl('ATBXGROUPECONF'));
// C.B 13/01/03 suite a modification de l'agl
// forcer complete a false
if TGroupeConf <> nil then TGroupeConf.Complete := false;

If (V_PGI.Superviseur) or (V_PGI.Controleur) then
 begin
  If (TGroupeConf=Nil) then exit;
  if (GetParamsoc('SO_AFTYPECONF')<>'AGR') then
   begin
    TGroupeConf.Visible := False;
    LTGroupeConf := THLabel(GetControl('T'+TGroupeConf.name));
    //NCX 10/01/02
    if LTGroupeConf <> Nil then LTGroupeConf.visible := False;
   end;
  exit;
 end;
//Confidentialité par groupe
if (GetParamsoc('SO_AFTYPECONF')='AGR') then
    begin
      If (TGroupeConf=Nil) then exit;
      //Récupération des groupes de confidentialité auxquel appartient l'utilisateur en cours.
      vargroupeConf := VH_GC.AfGereCritGroupeConf;
      if vargroupeConf = '' then exit;
      critere := ReadTokenPipe(vargroupeConf,';ComboPlus;');
      plus := varGroupeConf;
      //Redondance car control disabled mais permettra une gestion par le plus ou par la zone....
      TGroupeConf.text := critere;
      TGroupeConf.Plus := 'AND ('+plus +')';
      //
      TGroupeConf.enabled := False;
    end
else
 BEGIN
    If (TGroupeConf<>Nil) then
     begin
      TGroupeConf.Visible := False;
      THLabel(GetControl('T'+TGroupeConf.name)).visible := False;
     end;
    //Confidentialité par responsable
    if (GetParamsoc('SO_AFTYPECONF')='A01') then
        begin
        TconfLoc := ThEdit(GetControl('AFF_RESPONSABLE'+deux));
        //if (TconfLoc=nil) then
        //    TconfLoc := THEdit(GetControl('EAF_RESPONSABLE'+deux));
        if (TconfLoc=nil) then
            TconfLoc := THEdit(GetControl('ATBXRESPONSABLE'+deux));
        Valinit := VH_GC.RessourceUser;
        end
    else
    //Confidentialité par assistant libre
    For i:=1 to 3 do
         if (GetParamsoc('SO_AFTYPECONF')='AS'+IntToStr(i)) then
            begin
            TconfLoc := THEdit(GetControl('AFF_RESSOURCE'+IntToStr(i)+deux));
          //  if (TconfLoc=nil) then
           //     TconfLoc := THEdit(GetControl('EAF_RESSOURCE'+IntToStr(i)+deux));
            if (TconfLoc=nil) then
                TconfLoc := THEdit(GetControl('ATBXAFFRESSOURCE'+IntToStr(i)+deux));
            Valinit := VH_GC.RessourceUser;
            end;
    //on affecte à la zone de confidentialité sa valeur.
    If (TconfLoc<>Nil) then
       begin
        TconfLoc.text := Valinit;
        TconfLoc.enabled := False;
       end;
 END;

end;

procedure TOF_AFBASECODEAFFAIRE.bSelectAff1Click(Sender: TObject);
begin

  if assigned(GetControl('AFF_AFFAIRE0')) then
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('AFF_AFFAIRE0')),EditAff1,EditAff2,EditAff3,EditAff4)
  else if assigned(GetControl('AFFAIRE0')) then
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('AFFAIRE0')),EditAff1,EditAff2,EditAff3,EditAff4)
  else if assigned(GetControl('CH_CHANTIER0')) then
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('CH_CHANTIER0')),EditAff1,EditAff2,EditAff3,EditAff4)
  else
	SelectionAffaire (EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4);

end;


procedure TOF_AFBASECODEAFFAIRE.bSelectConClick(Sender: TObject);
begin
  EditAff0.Text := 'I';
	SelectionAffaire (EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4);
end;

procedure TOF_AFBASECODEAFFAIRE.bSelectAff2Click(Sender: TObject);
begin
//  EditAff0.Text := 'A';
//	SelectionAffaire (EditTiers_,EditAff_,EditAff0_,EditAff1_,EditAff2_,EditAff3_,EditAff4_);

  if assigned(GetControl('AFF_AFFAIRE0')) then
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('AFF_AFFAIRE0')),EditAff1,EditAff2,EditAff3,EditAff4)
  else if assigned(GetControl('AFFAIRE0')) then
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('AFFAIRE0')),EditAff1,EditAff2,EditAff3,EditAff4)
  Else if assigned(GetControl('CH_CHANTIER0')) then
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('CH_CHANTIER0')),EditAff1,EditAff2,EditAff3,EditAff4)
  else
    SelectionAffaire (EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4);

end;


procedure TOF_AFBASECODEAFFAIRE.SelectionAffaire(EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4 :THEdit);
var
Aff0,Aff1, Aff2, Aff3, Aff4,Affaire, VenteAchat : string;
bProposition : Boolean;
begin

  Aff0:='';
  Aff1:='';
  Aff2:='';
  Aff3:='';
  Aff4:='';
  Affaire :='';
  bProposition := False;
  VenteAchat:='';

  if THEdit(GetControl('GP_VENTEACHAT')) <> Nil then VenteAchat := GetControlText('GP_VENTEACHAT');

  if (EditAff0 <> Nil) then
    if EditAff0.Text = 'P' then bProposition := true;
  //
  if EditAff <> Nil then Affaire := EditAff.Text;

  if Ecran.Name = 'BTSAISIECONSO' then
    if EditAff0.text = 'W' then TheEtatAffaire := 'AFF,ECR,REA';

//il faut savoir ici avant l'appel de la recherche affaire si nous sommes sur un type d'affaire unique ou multiple !!!!
//afin de charger correctement le EditAff0 et gérer en aval le thcombobox de sélection du type d'affaire (Chantier, Contrat, Intervention)

  if GetAffaireEnteteSt(EditAff0,EditAff1,EditAff2,EditAff3,EditAff4,EditTiers, Affaire,bProposition,bchangeStatut,false,bChangeTiers,BcreatAff,VenteAchat,false,true,false,TheEtatAffaire ) then
  begin
  //--- Modif BRL 160604
  if (EditTiers <> Nil) and (VenteAchat <> 'VEN') then
    EditTiers.Text:='';

  //--- Decoupe du code affaire dans le cadre d'une appli BTP
  {$IFDEF BTP}
  BTPCodeAffaireDecoupe(Affaire,Aff0,Aff1,Aff2,Aff3,Aff4,taCreat, false);
  //end;
  {$ELSE}
  CodeAffaireDecoupe(Affaire,Aff0,Aff1,Aff2,Aff3,Aff4,taCreat, false);
  {$ENDIF}

  if EditAff <> Nil then EditAff.Text:= Affaire;
  if EditAff0 <> Nil then EditAff0.Text:= Aff0;

  EditAff1.Text:= Aff1;

                  // mcd 01/10/01 ajout test sur zones visibles, sinon sélection 'de à' ne fonctionne pas !!!
                  // mcd 17/05/02 ajout effacement zone si non géré, sinon, laisse ancienne valeur (bug complligne si change mission)
  if status = 'APP' then
  begin
    if GetParamSoc('SO_APPCO2Visible')  then EditAff2.Text:= Aff2 else EditAff2.text:='';
    if GetParamSoc('SO_APPCO3Visible')  then EditAff3.Text:= Aff3 else EditAff3.text:='';
  end else
  begin
    if GetParamSoc('SO_affCO2Visible')  then EditAff2.Text:= Aff2 else EditAff2.text:='';
    if GetParamSoc('SO_affCO3Visible')  then EditAff3.Text:= Aff3 else EditAff3.text:='';
  end;
  if (EditAff4 <> Nil) And (GetParamSoc('SO_AFFGESTIONAVENANT')) then EditAff4.Text:= Aff4;

  if (Ecran is TFMul) then TFMul(Ecran).ChercheClick;
end;
end;

procedure TOF_AFBASECODEAFFAIRE.bEffaceAff1Click(Sender: TObject);
var
  bFocus : boolean;
begin
  bFocus := false;
  // PL le 15/01/02 pour V575
  if (EditAff1_ <>nil) then
      EffaceAffaire (EditTiers_,EditAff_,EditAff0_,EditAff1_,EditAff2_,EditAff3_,EditAff4_);
  EffaceAffaire (EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4);

  if (Ecran is TFMul) then
     TFMul(Ecran).ChercheClick;

  if (Editaff0 <> Nil) then
  if EditAff0.CanFocus then
      BEGIN
      EditAff0.SetFocus;
      bFocus := true;
      END;

  if (Not (bFocus)) and (EditAff1.CanFocus) then EditAff1.SetFocus;

end;

procedure TOF_AFBASECODEAFFAIRE.bEffaceAff2Click(Sender: TObject);
var
  bFocus:boolean;
begin
  bFocus := false;
  EffaceAffaire (EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4);
  EffaceAffaire (EditTiers_,EditAff_,EditAff0_,EditAff1_,EditAff2_,EditAff3_,EditAff4_);

  if (Ecran is TFMul) then
     TFMul(Ecran).ChercheClick;

  if (EditAff0_ <> Nil) then
  if EditAff0_.CanFocus then
      BEGIN
      EditAff0_.SetFocus;
      bFocus := true;
      END;

  if Not (bFocus) then EditAff1_.SetFocus;
end;

procedure TOF_AFBASECODEAFFAIRE.EffaceAffaire(EditTiers,EditAff,EditAff0 ,EditAff1,EditAff2,EditAff3,EditAff4 :THEdit);
//Var bFocus : Boolean;
begin
  if (EditTiers<>nil) then
     if EditTiers.Enabled then EditTiers.Text:='';
  if (EditAff4<>nil) then
     EditAff4.Text:='';
  if (EditAff3<>nil) then
		 EditAff3.Text:='';
  if (EditAff2<>nil) then
	   EditAff2.Text:='';
  if (EditAff1<>nil) then
	   EditAff1.Text:='';

  {$IFNDEF BTP}
  if (Editaff0 <> Nil) then
      BEGIN
      EditAff0.Text:='';
  //    if EditAff0.CanFocus then
  //        BEGIN
  //        EditAff0.SetFocus;
  //        bFocus := true;
  //        END;
      END;
  {$ENDIF}
  if (EditAff <> Nil) then EditAff.Text:='';
  //if Not (bFocus) then EditAff1.SetFocus;

  //if (Ecran is TFMul) then
  //   TFMul(Ecran).ChercheClick;
end;

procedure TOF_AFBASECODEAFFAIRE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
// Le champ AFFAIRE est obligatoire dans la fiche
// Il sera renseigné par la fct appelante si accès direct à une affaire
// Ce code ne doit autrement jamais être renseigné par le prog

Aff:=THEdit(GetControl('ACT_AFFAIRE'));
Aff0:=THEdit(GetControl('ACT_AFFAIRE0'));
Aff1:=THEdit(GetControl('ACT_AFFAIRE1'));
Aff2:=THEdit(GetControl('ACT_AFFAIRE2'));
Aff3:=THEdit(GetControl('ACT_AFFAIRE3'));
Aff4:=THEdit(GetControl('ACT_AVENANT'));
Tiers:=THEdit(GetControl('ACT_TIERS'));

Aff_:=THEdit(GetControl('ACT_AFFAIRE_'));
Aff0_:=THEdit(GetControl('ACT_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('ACT_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('ACT_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('ACT_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('ACT_AVENANT_'));

Tiers_:=THEdit(GetControl('ACT_TIERS_'));

end;

procedure TOF_AFBASECODEAFFAIRE.TypeAppelChargeCle(var AvActionFiche:TActionFiche);
begin
AvActionFiche := taCreat;
end;

procedure TOF_AFBASECODEAFFAIRE.bSelectAffMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var Tiers					 : string;
		Aff0					 : string;
		Aff1					 : string;
		Aff2					 : string;
		Aff3					 : string;
		Aff4					 : string;
		Affaire 			 : string;

		ChampDestTiers :  THEdit;
		ChampDestAff	 :  THEdit;
		ChampDestAff0	 :  THEdit;
		ChampDestAff1	 :  THEdit;
		ChampDestAff2	 :  THEdit;
		ChampDestAff3	 :  THEdit;
		ChampDestAff4  :  THEdit;

begin

  if (Shift<>[ssCtrl,ssLeft]) then exit;

  Aff0 :='';
  Aff1 :='';
  Aff2 :='';
  Aff3 :='';
  Aff4 :='';
  Affaire :='';

  if (TToolBarButton97(Sender).Name = 'BSELECTAFF1') then
      begin
      ChampDestTiers := EditTiers;
      ChampDestAff   := EditAff;
      ChampDestAff0  := EditAff0;
      ChampDestAff1  := EditAff1;
      ChampDestAff2  := EditAff2;
      ChampDestAff3  := EditAff3;
      ChampDestAff4  := EditAff4;
      if (EditTiers_<>nil) then Tiers 	:= EditTiers_.Text;
      if (EditAff0_<>nil)  then Aff0		:= EditAff0_.Text;
      if (EditAff1_<>nil)  then Aff1 		:= EditAff1_.Text;
      if (EditAff2_<>nil)  then Aff2 		:= EditAff2_.Text;
      if (EditAff3_<>nil)  then Aff3 		:= EditAff3_.Text;
      if (EditAff4_<>nil)  then Aff4 		:= EditAff4_.Text;
      if (EditAff_<>nil)   then Affaire := EditAff_.Text;
      end
  else if (TToolBarButton97(Sender).Name = 'BSELECTAFF2') then
      begin
      ChampDestTiers := EditTiers_;
      ChampDestAff   := EditAff_;
      ChampDestAff0  := EditAff0_;
      ChampDestAff1  := EditAff1_;
      ChampDestAff2  := EditAff2_;
      ChampDestAff3  := EditAff3_;
      ChampDestAff4  := EditAff4_;
      if (EditTiers<>nil) then Tiers 	 := EditTiers.Text;
      if (EditAff0<>nil)  then Aff0 	 := EditAff0.Text;
      if (EditAff1<>nil)  then Aff1 	 := EditAff1.Text;
      if (EditAff2<>nil)  then Aff2 	 := EditAff2.Text;
      if (EditAff3<>nil)  then Aff3 	 := EditAff3.Text;
      if (EditAff4<>nil)  then Aff4 	 := EditAff4.Text;
      if (EditAff<>nil)   then Affaire := EditAff.Text;
      end
  else
      exit;

  if (ChampDestTiers <> nil) then ChampDestTiers.Text := Tiers;
  if (ChampDestAff   <> Nil) then ChampDestAff.Text:= Affaire;
  if (ChampDestAff0  <> Nil) then ChampDestAff0.Text:= Aff0;
  if (ChampDestAff1  <> Nil) then ChampDestAff1.Text:= Aff1;
  if (ChampDestAff2  <> Nil) and  GetParamSoc('SO_affCO2Visible')     then ChampDestAff2.Text:= Aff2;
  if (ChampDestAff3  <> Nil) and  GetParamSoc('SO_affCO3Visible')     then ChampDestAff3.Text:= Aff3;
  if (ChampDestAff4  <> Nil) and  GetParamSoc('SO_AFFGESTIONAVENANT') then ChampDestAff4.Text:= Aff4;

	Abort;

end;

procedure TOF_AFBASECODEAFFAIRE.OnLoad;
begin
  inherited;
  //
  //NCX 13/11/2001 Confidentialité affaires
  //
  GereChampConfidentialite(Tconf,'');
  GereChampConfidentialite(Tconf_,'_');

  {$IFDEF CCS3}
  if (Getcontrol('AFF_MODELE')<> Nil) then setcontrolVisible('AFF_MODELE',false);
  {$ENDIF}
end;




Initialization
registerclasses([TOF_AFBASECODEAFFAIRE]);


end.
