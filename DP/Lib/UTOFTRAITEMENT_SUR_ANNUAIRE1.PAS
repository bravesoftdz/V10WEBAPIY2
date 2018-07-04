{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : TRAITEMENT_SUR_ANNUAIRE ()
               : sources des fct sans aucun line avec GC et GI
Mots clefs ... : TOF;TRAITEMENT_SUR_ANNUAIRE
*****************************************************************}

//mcd 14/02/2006 diverses modif pour passer en GUID au lieu de codeper

Unit UTOFTRAITEMENT_SUR_ANNUAIRE1 ;

Interface

Uses
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     FE_Main, Mul,{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} db,HDB,
{$ENDIF}
     Utob,  HStatus, forms, UTOF,M3fp, Classes,
     Hctrls, Hent1,HMsgBox,HTB97,HQry,ComCtrls, DPTofAnnuSel,
     SysUtils,Vierge, StdCtrls, Windows, AGLInit;

Type
  // fiche AFCTRL_ANNTIERS : mul pour choisir clt sans lien annauire
  TOF_AFCtrl_Anntiers = Class (TOF)
     procedure Ctrl_anntiers;
     procedure Onload;    override;
     procedure OnArgument(stArgument : String ); override ;
     procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
     Where, WhereOri: String;
     procedure SetAnn;
     procedure SetAllAnn;
  end ;
  // fiche TIERS_ANN : pour sélection rapide d'une fiche annuaire avec un client
  TOF_Tiers_Ann_Align = Class (TOF)
     Procedure OnArgument(stArgument : String ); override ;
     procedure SelAnn;
  end ;

  // fiche AFDOUBLONS_ANNTIE : liste des fiche annuaires avec même
  /// lien client
  TOF_AFDOUBLONS_ANNTIERS = Class (TOF)
     procedure Doublons_anntiers;
     Procedure OnArgument(stArgument : String ); override ;
     procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
     procedure SetDoublAnn;
  end ;

  // fiche AFDOUBLONS_ANNU : liste pour sélection des annuaires en double
  TOF_AFDOUBLONS_ANNU = Class (TOF)
     Procedure OnArgument(stArgument : String ); override ;
     Procedure OnLoad; override ;
     procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
     Procedure BOuvrirOnClick(sender : Tobject);
  end ;

  //fiche AFSUPRDOUBLONANNU pour sélection fiche annuaire à conserver / à supprimer
  TOF_AFSUPRDOUBLONANNU = Class (TOF)
     Procedure OnArgument(stArgument : String ); override ;
  private
     Procedure ChargeEnreg( st : string);
     Procedure SupprimeOnClick(Sender : TObject);
     Procedure ForceOnClick(Sender : TObject);
     Function ChercheLiens(num, GuidPer : string) : integer;
     function GetControlChecked(nomzone: String): Boolean;
  end ;

procedure AGLDoublons_Anntiers( parms: array of variant; nb: integer );
procedure AGLCtrl_Anntiers( parms: array of variant; nb: integer );

Function AFLanceFiche_TiersSansAnn(Argument:string):variant;
Procedure AFLanceFiche_SuprDoublAnn(Argument:string);
Procedure AFLanceFiche_MulAnnuTiersFaux;
Procedure AFLanceFiche_MulAnnTiersDouble;
Procedure AFLanceFiche_MulDoublonsAnn;

Implementation

uses AnnOutils, AnnDedoub;
 const
  TexteMsg: array [1..5] of string = (
    {1}   'Impossible de mettre à jour l''annuaire '
    {2} , 'Impossible de mettre à jour tous les annuaires'
    {3} , 'Vous devez sélectionner deux fiches obligatoirement.'
    {4} , 'Sélection doublons'
    {5} , 'Liens plaquette IFU.'
            ) ;

//******************************************************************************
//******************** Gestion Alignement Annuaire/client **********************
// *** affiche client sans lien annuaire et permet mise à jour rapide
//******************************************************************************

procedure TOF_AFCtrl_Anntiers.OnArgument(stArgument : String );
begin
  Ecran.OnKeyDown := Form_OnKeyDown;
  Whereori := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
  inherited;
end;

procedure TOF_AFCtrl_Anntiers.Onload;
var twc : string;
begin
  TWC := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
  Where := ')';
  // il faut récupérer la sélection éventuelle faite pour utilisation sur le select all
  if Length(Twc) > Length(whereori) then
    begin
    Where := Copy (Twc, Length(whereori), Length(Twc));
    end;
end;

procedure TOF_AFCtrl_Anntiers.Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 case Key of
  65 : if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) then
        if(TFMul(Ecran).FListe.AllSelected) then
         TFMul(Ecran).FListe.AllSelected:=False
        else
         TFMul(Ecran).FListe.AllSelected:=True;

  VK_F10 : begin
            AGLCtrl_AnnTiers ([LongInt(Ecran)], 1);
            AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
           end;
  else
   TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 end;
end;

procedure TOF_AFCtrl_Anntiers.Ctrl_AnnTiers;
var i: integer;
begin
  with TFMul(Ecran) do
    begin
    if FListe.AllSelected then
      BEGIN
      if Transactions(SetAllAnn,3) <> oeOK then PGIBox(TexteMsg[1]+Q.FindField('ANN_GUIDPER').AsString, Caption);
      FListe.AllSelected := false;
      TToolBarButton97(GetControl('bSelectAll')).Down:=false;
      END
    else
      BEGIN
      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
        BEGIN
        FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row-1);
{$ELSE}
{$ENDIF}
        if Transactions(SetAnn,3) <> oeOK then PGIBox(TexteMsg[1]+Q.FindField('ANN_NOMPER').AsString, Caption);
        MoveCur(False);
        END;
      FListe.ClearSelected;
      FiniMove;
      END;
    ChercheClick;
    end;
end;


// procédure appelée pour mettre à jour les champs dans annuaire sur enrgt sélectionné
procedure TOF_AFCtrl_Anntiers.SetAnn;
var Guidper : string;
begin
  Guidper:='';
  with TFMul(Ecran) do
    begin
    Guidper:= Q.FindField('ANN_GUIDPER').AsString;
    end;
  ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" '+
             'WHERE '+
             'ANN_GUIDPER="'+Guidper+'"');
end ;


// Procédure appelée par Transactions pour mettre à jour tout annuaire
// en une seule requête
procedure TOF_AFCtrl_Anntiers.SetAllAnn;
begin
  ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" '+
          //mcd 17/05/05 champ supprime      ',ANN_AUXILIAIRE=""'+
             ' where (ann_tiers<>"" and   (select t_tiers from tiers where t_tiers=ann_tiers) is null '+
             Where);
end;


procedure AGLCtrl_Anntiers( parms: array of variant; nb: integer );
var F : TForm;
    LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (Latof is TOF_AFCtrl_Anntiers) then TOF_AFCtrl_Anntiers(LaTof).Ctrl_Anntiers else exit;
end;


//******************************************************************************
//************* Gestion des annuaires ayant un lien sur le même client**********
//******************************************************************************
procedure TOF_AFDOublons_Anntiers.OnArgument(stArgument : String );
begin
 Ecran.OnKeyDown := Form_OnKeyDown;
 inherited;
end;

procedure TOF_AFDOublons_Anntiers.Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 case Key of
  65 : if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) then
        if(TFMul(Ecran).FListe.AllSelected) then
         TFMul(Ecran).FListe.AllSelected:=False
        else
         TFMul(Ecran).FListe.AllSelected:=True;

  VK_F10 : begin
            AGLDoublons_Anntiers ([LongInt(Ecran)], 1);
            AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
           end;
  else
       TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 end;
end;

procedure TOF_AFDOublons_Anntiers.Doublons_AnnTiers;
var  i: integer;
begin
  with TFMul(Ecran) do
    begin
    if FListe.AllSelected then
      BEGIN
      // cas impossible...il ne faut sélectionner qu'un client sur 2 ....
      PGIBox(TexteMsg[1], Caption);
      END
    else
      BEGIN
      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
        BEGIN
        FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row-1);
{$ELSE}
{$ENDIF}
        if Transactions(SetDoublAnn,3) <> oeOK then PGIBox(TexteMsg[1]+Q.FindField('ANN_NOMPER').AsString, Caption);
        MoveCur(False);
        END;
      FListe.ClearSelected;
      FiniMove;
      END;
   ChercheClick;
   end;
end;


// procédure appelée pour mettre à jour les champs dans annuaire sur enrgt sélectionné
procedure TOF_AFDOublons_Anntiers.SetDoublAnn;
var Guidper : string;
begin
  with TFMul(Ecran) do
    begin
    Guidper:= Q.FindField('ANN_GUIDPER').AsString;
    end;
  ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" '+
            //mcd 17/05/05 champ supprimé  ',ANN_AUXILIAIRE="" '+
             'WHERE '+
             'ANN_GUIDPER="'+Guidper+'"');
end ;


procedure AGLDOublons_Anntiers( parms: array of variant; nb: integer );
var F : TForm;
    LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (Latof is TOF_AFDOublons_Anntiers) then TOF_AFDOublons_Anntiers(LaTof).Doublons_Anntiers else exit;
end;


//******************************************************************************
//********************* Gestion Doublons Annuaire******************************
//******************************************************************************

procedure TOF_AFDoublons_Annu.BOuvrirOnClick(sender: Tobject);
var i : integer;
    parm : string;
    QQ : THquery;
begin
  parm :='';
  with TFMul(Ecran) do
    begin
    If FListe.NbSelected <> 2 then PgiInfo(TexteMsg[3],TexteMsg[4])
    else
      begin
      QQ := TFMul(Ecran).Q;
      For i:=0 to FListe.NbSelected-1 do
        begin
        FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
        QQ.TQ.Seek(Fliste.Row - 1) ;  //mcd 09/03/2006
{$ENDIF}
        parm := Parm+QQ.FindField('ANN_GUIDPER').asString+';';
        end;
      AFLanceFiche_SuprDoublAnn(parm);
      FListe.ClearSelected;
      end;
    ChercheClick;
    end;
end;

procedure TOF_AFDoublons_Annu.OnLoad;
var suf, st : string;
begin
  inherited;
  IF ThRadioGroup(GetControl('RGP')).Value <> 'AUCUN' then
    begin
    suf := ThRadioGroup(GetControl('RGP')).Value;
    st := 'ANN_'+suf+' in (select ann_'+suf+' from annuaire group by ann_'+suf+' having count(*) > 1)';
    // $$$ JP 04/09/06: pour D7/Unicode: ne plus faire ce genre de cast... ThEdit(GetControl('XX_WHERE')).text := st;
    SetControlText ('XX_WHERE', st);

    // pas besoin de tri, le fait tout seul ...
    // sinon, il faut ajouter un champ order by dans la fiche
    //sinon pas oK dans le xx_where st:=  'order by ann_'+suf;
    end
  else
      // $$$ JP 04/09/06: pour D7/Unicode: ne plus faire ce genre de cast... ThEdit(GetControl('XX_WHERE')).text := '';
      SetControlText ('XX_WHERE', '');
end;

procedure TOF_AFDOublons_Annu.OnArgument(stArgument : String );
begin
 Ecran.OnKeyDown := Form_OnKeyDown;
 TToolbarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirOnClick;
 inherited;
end;

procedure TOF_AFDOublons_Annu.Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 case Key of
  65 : if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) then
        if(TFMul(Ecran).FListe.AllSelected) then
         TFMul(Ecran).FListe.AllSelected:=False
        else
         TFMul(Ecran).FListe.AllSelected:=True;

  VK_F10 : BOuvrirOnclick (nil);

  else
       TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 end;
end;

{ TOF_AFSUPRDOUBLONANNU }
//******************************************************************************
//********************* Fiche pour choix annauire à supprimer*******************
//******************************************************************************

procedure TOF_AFSUPRDOUBLONANNU.OnArgument(stArgument: String);
begin
  ChargeEnreg(StArgument); // affiche les données + les labels voulus
  TToolbarButton97(GetControl('BSUPPR1')).OnClick := SupprimeOnClick;
  TToolbarButton97(GetControl('BSUPPR2')).OnClick := SupprimeOnClick;
  TToolbarButton97(GetControl('FORCE1')).OnClick := ForceOnClick;
  TToolbarButton97(GetControl('FORCE2')).OnClick := ForceOnClick;
  inherited;
end;


procedure TOF_AFSUPRDOUBLONANNU.ChargeEnreg(st: string);
var QQ : Tquery;
    i,ret : integer;
    Listechamp,ListechampTemp,champ, champTemp,nomzone, Guidper : string;
    zone : thEdit;
begin
  ListeChamp := 'ANN_ALCP,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,ANN_ALVILLE,'+
                //mcd 17/05/05 champ supprime    'ANN_AUXILIAIRE,ANN_CLESIRET,ANN_CODENAF,ANN_CODEPER,'+
                'ANN_CLESIRET,ANN_CODENAF,ANN_GUIDPER,'+
                'ANN_EMAIL,ANN_FAX,ANN_FORME,ANN_FORMEGEN,'+
                'ANN_MINITEL,ANN_MOISCLOTURE,ANN_NATIONALITE,ANN_NOM1,ANN_NOM2,'+
                'ANN_NOM3,ANN_NOMPER,ANN_PAYS,ANN_RCS,ANN_SIREN,ANN_SITEWEB,ANN_TEL1,ANN_TIERS,ANN_TYPEPER';
  For i := 1 to 2 do
    begin
    Guidper := ReadTokenSt(St);
    QQ := OpenSQL('SELECT '+ListeChamp+' FROM ANNUAIRE WHERE ANN_GUIDPER="'+Guidper+'"', True);
    ListechampTemp := Listechamp;
    While ListeChampTemp<>'' do
      begin
      Champ := ReadTokenPipe(ListechampTemp,',');
      ChampTemp := ReadTokenPipe(Champ,'_');
      NomZone := 'AN'+IntToStr(i)+'_'+Champ;
      zone := Thedit(GetControl(Nomzone));
      If zone <> Nil then SetControlText(Nomzone,QQ.FindField('ANN_'+champ).asString);
      end;
    Ferme(QQ);
    ret := ChercheLiens(IntToStr(i),GuidPer);
    // si existence de liens rhédibitoires pour la suppression,
    // on désactive le bouton
    if (ret>0) and ( GetControlChecked('CBJURIDIQUE'+IntToStr(i))
     or GetControlChecked( 'CBLIENS'+IntToStr(i))
     or GetControlChecked( 'CBIFU'+IntToStr(i))
     or GetControlChecked( 'CBDOSSIER'+IntToStr(i)) ) then
      SetControlEnabled ('BSUPPR'+ IntToStr(i),False);
    end;
  SetControlTExt ('AN1_GUIDPER1',GetControltext('AN1_GUIDPER'));
  SetControlTExt ('AN2_GUIDPER1',GetControltext('AN2_GUIDPER'));
  SetControlVisible('FORCE1',true);
  setControlVisible('FORCE2',true);
end;


Function TOF_AFSUPRDOUBLONANNU.ChercheLiens(num, Guidper : string) : integer;
var nbliens : integer;
    Q : TQuery;
    PerSup: TPerSup;
    Msg : String;
begin
  Result := 0;
  PerSup := TPerSup.Create;
  PerSup.GuidPer := Guidper;
  PerSup.Bavard := False;

  // cherche si liens juridique (#### manque évènements)
  if (PerSup.TestJuriDosPer(Msg)) or (PerSup.TestJuriLienPer(Msg))
  or (PerSup.TestJuriDosInfoPer(Msg)) or (PerSup.TestJuriGroupeSocPer(Msg)) then
    begin
    SetControlChecked('CBJURIDIQUE'+num, True);
    SetControlCaption('LJURIDIQUE'+num, 'Cliquer sur <+ d''info>');
    SetControlVisible('LJURIDIQUE'+num, True);
    Result := Result + 1 ;
    end;

  // cherche si liens DP
  Q := OpenSql('SELECT COUNT(ANL_FONCTION) AS COMPTE FROM ANNULIEN WHERE ANL_GUIDPER="'+guidper+'"', True);
  If not Q.eof then
    begin
    nbliens := Q.FindField('COMPTE').AsInteger;
    If nbliens<>0 then
      begin
      SetControlChecked('CBLIENS'+num, True);
      SetControlCaption('LLIENS'+num, IntToStr(nbliens)+' lien(s)');
      SetControlVisible('LLIENS'+num, True);
      Result := Result + 1 ;
      end;
     end;
  Ferme(Q);

  // existence de dossiers de production (NPC avec dossiers partis) 
  Q := OpenSql('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_GUIDPER="'+Guidper+'"', True);
  If not Q.eof then
    begin
    SetControlChecked('CBDOSSIER'+num, True);
    SetControlCaption('LDOSSIER'+num, 'numéro : ' +Q.findfield('DOS_NODOSSIER').asString);
    SetControlVisible('LDOSSIER'+num, True);
    Result := Result + 1 ;
    end;
  Ferme(Q);

  // liens dans plaquettes (en fait, seule la plaquette IFU a des tables dans
  // la base commune, donc conclusion explicite...)
  if PerSup.TestLienPlaq(Msg) then
    begin
    SetControlChecked('CBIFU'+num, True);
    SetControlCaption('LIFU'+num,TexteMsg[5]);
    SetControlVisible('LIFU', True);
    end;

  // liens tiers compta
  If GetControlText('AN'+num+'_TIERS')<>'' then
    begin
    SetControlChecked('CBTIERS'+num, True);
    SetControlVisible('AN'+num+'_TIERS', True);
    Result := Result + 1 ;
    end;

  PerSup.Free;
END;


procedure TOF_AFSUPRDOUBLONANNU.SupprimeOnClick(Sender: TObject);
var nombouton, nobouton : String;
    GuidpersonneDP: String;
begin
  nombouton := TToolBarButton97(sender).name;
  // dernier carac
  nobouton := copy(nombouton,length(nombouton),1);
  GuidpersonneDP := GetControlText('AN'+nobouton+'_GUIDPER');
  if GuidpersonneDP<>'' then SupprimeEnregAnnu(GuidpersonneDP);
  Ecran.close;
end;


procedure TOF_AFSUPRDOUBLONANNU.ForceOnClick(Sender: TObject);
var nombouton, nobouton : string;
begin
  nombouton := TToolBarButton97(sender).name;
  nobouton := Copy(nombouton, Length(nombouton), 1);
  FicheAssistantDedoublonnage(TFVierge(Ecran), nobouton);
end;


function TOF_AFSUPRDOUBLONANNU.GetControlChecked(nomzone: String): Boolean;
var c: TComponent;
begin
  Result := False;
  c := GetControl(nomzone);
  if c is TCheckBox then
    Result := TCheckBox(c).Checked
  else if c is TRadioButton then
    Result := TRadioButton(c).Checked;
end;


//******************************************************************************
//********************* Ecran pour affecter annuaire ***************************
// fiche Tiers_ann : pour sélection rapide d'une fiche annuaire avec un client
//******************************************************************************

procedure TOF_Tiers_ann_Align.OnArgument(stArgument : String );
var Q : TQuery;
begin
  inherited;
  SetControlText('T_TIERS',stArgument);
  SetControlText('T_TIERS1',stArgument);
  Q := OpenSql ('Select T_PAYS,T_LIBELLE,T_NATIONALITE,T_LANGUE,T_DEVISE,'
    +'T_FAX,T_MOISCLOTURE,T_ABREGE,T_RVA,T_TELEPHONE,T_TELEX,T_APE,'
    +'T_SIRET,T_LIBELLE,T_PRENOM,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,'
    +'T_VILLE,T_CODEPOSTAL,T_FORMEJURIDIQUE from tiers where t_tiers="'
    + stArgument+'"', True);
  if not Q.Eof then
    begin
    SetControlText('T_PAYS',Q.Findfield('T_PAYS').AsString);
    SetControlText('T_NATIONALITE',Q.Findfield('T_NATIONALITE').AsString);
    SetControlText('T_LANGUE',Q.Findfield('T_LANGUE').AsString);
    SetControlText('T_DEVISE',Q.Findfield('T_DEVISE').AsString);
    SetControlText('T_FAX',Q.Findfield('T_FAX').AsString);
    SetControlText('T_MOISCLOTURE',InttoStr(Q.Findfield('T_MOISCLOTURE').AsInteger));
    SetControlText('T_ABREGE',Q.Findfield('T_ABREGE').AsString);
    SetControlText('T_RVA',Q.Findfield('T_RVA').AsString);
    SetControlText('T_TELEPHONE',Q.Findfield('T_TELEPHONE').AsString);
    SetControlText('T_TELEX',Q.Findfield('T_TELEX').AsString);
    SetControlText('T_APE',Q.Findfield('T_APE').AsString);
    SetControlText('T_SIRET',Q.Findfield('T_SIRET').AsString);
    SetControlText('T_LIBELLE',Q.Findfield('T_LIBELLE').AsString);
    SetControlText('T_PRENOM',Q.Findfield('T_PRENOM').AsString);
    SetControlText('T_ADRESSE1',Q.Findfield('T_ADRESSE1').AsString);
    SetControlText('T_ADRESSE2',Q.Findfield('T_ADRESSE2').AsString);
    SetControlText('T_ADRESSE3',Q.Findfield('T_ADRESSE3').AsString);
    SetControlText('T_VILLE',Q.Findfield('T_VILLE').AsString);
    SetControlText('T_CODEPOSTAL',Q.Findfield('T_CODEPOSTAL').AsString);
    SetControlText('T_FORMEJURIDIQUE',Q.Findfield('T_FORMEJURIDIQUE').AsString);
    end;
  Ferme (Q);
end;


procedure TOF_Tiers_ann_Align.SelAnn;
var Q : TQuery;
    Guidper:string;
begin
//  GuidPer:=AglLanceFiche('YY', 'ANNUAIRE_SEL','ANN_NOMPER='+GetControlText('T_ABREGE'),'','TIERS');
  GuidPer:=LancerAnnuSel ('ANN_NOMPER='+GetControlText('T_ABREGE'),'','TIERS');

  if (Guidper <>'') then
    begin
    SetControlText('ANN_GUIDPER',Guidper);
    SetControlText('ANN_GUIDPER1',Guidper);
    Q := OpenSql ('Select ANN_PAYS,ANN_NOM1,ANN_NATIONALITE,ANN_LANGUE,ANN_DEVISE,'
      +'ANN_FAX,ANN_MOISCLOTURE,ANN_NOMPER,ANN_SITEWEB,ANN_TEL1,ANN_MINITEL,ANN_CODENAF,'
      +'ANN_SIREN,ANN_CLESIRET,ANN_NOM1,ANN_NOM2,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,'
      +'ANN_ALVILLE,ANN_ALCP,ANN_FORME from Annuaire where ANN_GUIDPER="'
      + Guidper +'"', True);
    if not Q.Eof then
      begin
      SetControlText('ANN_PAYS',Q.Findfield('ANN_PAYS').AsString);
      SetControlText('ANN_NATIONALITE',Q.Findfield('ANN_NATIONALITE').AsString);
      SetControlText('ANN_LANGUE',Q.Findfield('ANN_LANGUE').AsString);
      SetControlText('ANN_DEVISE',Q.Findfield('ANN_DEVISE').AsString);
      SetControlText('ANN_FAX',Q.Findfield('ANN_FAX').AsString);
      SetControlText('ANN_MOISCLOTURE',Q.Findfield('ANN_MOISCLOTURE').AsString);
      SetControlText('ANN_NOMPER',Q.Findfield('ANN_NOMPER').AsString);
      SetControlText('ANN_SITEWEB',Q.Findfield('ANN_SITEWEB').AsString);
      SetControlText('ANN_TEL1',Q.Findfield('ANN_TEL1').AsString);
      SetControlText('ANN_MINITEL',Q.Findfield('ANN_MINITEL').AsString);
      SetControlText('ANN_CODENAF',Q.Findfield('ANN_CODENAF').AsString);
      SetControlText('ANN_SIREN',Q.Findfield('ANN_SIREN').AsString);
      SetControlText('ANN_CLESIRET',Q.Findfield('ANN_CLESIRET').AsString);
      SetControlText('ANN_NOM1',Q.Findfield('ANN_NOM1').AsString);
      SetControlText('ANN_NOM2',Q.Findfield('ANN_NOM2').AsString);
      SetControlText('ANN_ALRUE1',Q.Findfield('ANN_ALRUE1').AsString);
      SetControlText('ANN_ALRUE2',Q.Findfield('ANN_ALRUE2').AsString);
      SetControlText('ANN_ALRUE3',Q.Findfield('ANN_ALRUE3').AsString);
      SetControlText('ANN_ALVILLE',Q.Findfield('ANN_ALVILLE').AsString);
      SetControlText('ANN_ALCP',Q.Findfield('ANN_ALCP').AsString);
      SetControlText('ANN_FORME',Q.Findfield('ANN_FORME').AsString);
      end;
    Ferme(Q);
    end;
end;


procedure AGLSelAnn (parms: array of variant; nb: integer );
var F : TForm;
    LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFVierge) then Latof:=TFVierge(F).Latof else exit;
  If (LaTof is TOF_Tiers_ann_Align) then TOF_Tiers_ann_Align(LaTof).SelAnn else exit;
end;

 Procedure AFLanceFiche_MulAnnuTiersFaux;
begin
AGLLanceFiche('AFF','AFCTRL_ANNTIERS','','','');
end;
Procedure AFLanceFiche_MulAnnTiersDOuble;
begin
AGLLanceFiche('AFF','AFDOUBLONS_ANNTIE','','','');
end;
Procedure AFLanceFiche_MulDoublonsAnn;
begin
AGLLanceFiche('AFF','AFDOUBLONS_ANNU','','','');
end;
Function AFLanceFiche_TiersSansAnn(Argument:string):variant;
begin
result:=AGLLanceFiche('AFF','TIERS_ANN','','',argument);
end;
Procedure AFLanceFiche_SuprDoublAnn(Argument:string);
begin
AGLLanceFiche('AFF','AFSUPRDOUBLONANNU','','',argument);
end;

Initialization
RegisterClasses ( [TOF_AFCtrl_AnnTiers,TOF_AFDoublons_AnnTiers,TOF_AFDoublons_Annu,TOF_AFSUPRDOUBLONANNU,Tof_Tiers_Ann_Align] ) ;
RegisterAglProc( 'Doublons_Anntiers',True,0,AGLDoublons_Anntiers);
RegisterAglProc( 'SelAnn',True,0,AGLSelAnn);
RegisterAglProc( 'Ctrl_Anntiers',True,0,AGLCtrl_Anntiers);
end.
