{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/03/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BETATPREPPAIE ()
Mots clefs ... : TOF;BETATPREPPAIE
*****************************************************************}
Unit BETATPREPPAIE_TOF ;

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
     UTOb,
     UTOF,
     hPDFPrev,
     hPDFViewer,
     //EntGC,
     HTB97,
     Hpanel,
     //uEntCommun,
     //UtilTOBPiece,
     //uRecupSQLModele,
     UtilsEtat;
     // paramsoc,
     // Ent1,
     // UCotraitance;

Type
  TOF_BETATPREPPAIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    //
    TobCONSO      : TOB;
    //
    TobEditions   : TOB;
    //
    BValider      : TToolbarButton97;
    //
    Ressource     : THEdit;
    Ressource1    : THEdit;
    DateDeb       : THEdit;
    DateFin       : THEdit;
    //
    FETAT         : THValComboBox;
    TEtat         : ThLabel;
    //
    Idef          : Integer;
    //
    OptionEdition : TOptionEdition;
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;
    //
    BParamEtat    : TToolBarButton97;
    //
    ChkApercu     : TCheckBox;
    ChkReduire    : TCheckBox;
    ChkCouleur    : TCheckBox;
    ChkDefinitif  : TCheckBox;
    //
    procedure OnClickValide(Sender : Tobject);
    procedure ChargeEtat;
    procedure ControleChamp(Champ, Valeur: String);
    procedure CreateTobs;
    procedure GetObjects;
    procedure InitScreenObject;
    procedure ParamEtat(Sender: TOBJect);
    procedure ReleaseTob;
    procedure SetScreenEvents;
    procedure VisibleScreenObject;
    procedure OnClickApercu(Sender: Tobject);
    procedure OnClickReduire(Sender: Tobject);
    procedure OnChangeEtat(Sender: Tobject);
    procedure OnClickDefinitif(Sender: Tobject);
    procedure OnClickCouleur(Sender: Tobject);
    procedure EditeTout;
    procedure ChargementLigneEdition(TobEdition : TOB);
    procedure CreateLigneEdition(TOBEdition,Toblig: TOB; iMonth: word);


  end ;

Implementation

procedure TOF_BETATPREPPAIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BETATPREPPAIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BETATPREPPAIE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BETATPREPPAIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BETATPREPPAIE.OnArgument (stArgument : String ) ;
var Critere  : string;
    ValMul   : string;
    ChampMul : string;
    x        : integer;
begin
  Inherited ;

  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    valMul := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
      end
      else
        ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
    end;
  until Critere='';

  GetObjects;
  SetScreenEvents;
  VisibleScreenObject;
  InitScreenObject;

  //chargement de l'état en fonction du type de paiement : si direct Attestation, si regroupement eclatement
  ChargeEtat;

  VisibleScreenObject;

end ;

procedure TOF_BETATPREPPAIE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BETATPREPPAIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BETATPREPPAIE.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BETATPREPPAIE.GetObjects;
begin

  if assigned(GetControl('RESSOURCE'))  then Ressource := THEdit(Ecran.FindComponent('Ressource'));
  if assigned(GetControl('RESSOURCE1')) then Ressource1:= THEdit(Ecran.FindComponent('Ressource1'));
  if assigned(GetControl('DATEEDIT'))   then DateDeb   := THEdit(Ecran.FindComponent('DATEEDIT'));
  if assigned(GetControl('DATEEDIT_'))  then DateFin   := THEdit(Ecran.FindComponent('DATEEDIT_'));

  if Assigned(GetControl('BParamEtat')) then BParamEtat  := TToolbarButton97(ecran.FindComponent('BParamEtat'));

  if Assigned(GetControl('BValider'))   then BValider    := TToolbarButton97(ecran.FindComponent('BValider'));

  if Assigned(GetControl('fApercu'))    then ChkApercu   := TCheckBox(Ecran.FindComponent('fApercu'));
  if Assigned(GetControl('FReduire'))   then ChkReduire  := TCheckBox(Ecran.FindComponent('fReduire'));
  if Assigned(GetControl('fCouleur'))   then ChkCouleur  := TCheckBox(Ecran.FindComponent('fCouleur'));
  if Assigned(GetControl('fDefinitif')) then ChkDefinitif:= TCheckBox(Ecran.FindComponent('fDefinitif'));

  if Assigned(GetControl('fEtat'))      then FEtat       := ThValComboBox(ecran.FindComponent('FEtat'));
  if assigned(Getcontrol('TEtat'))      then TEtat       := ThLabel(ecran.FindComponent('TEtat'));

end;

procedure TOF_BETATPREPPAIE.SetScreenEvents;
begin

  if assigned(BValider)     then BValider.OnClick     := OnClickValide;

  if assigned(ChkApercu)    then ChkApercu.OnClick    := OnClickApercu;
  if assigned(ChkReduire)   then ChkReduire.OnClick   := OnClickReduire;
  if assigned(ChkCouleur)   then ChkCouleur.OnClick   := OnClickCouleur;
  if Assigned(ChkDefinitif) then ChkDefinitif.OnClick := OnClickDefinitif;

  if Assigned(BParamEtat)   then BParamEtat.OnClick   := ParamEtat;

  if assigned(FETAT)        then FETAT.OnChange       := OnChangeEtat;

end;

procedure TOF_BETATPREPPAIE.VisibleScreenObject;
begin

  if Fetat <> nil then
  begin
    if fetat.Items.Count > 1 then
    begin
      Fetat.Visible := true;
      BparamEtat.Visible := true;
    end;
  end;

end;

procedure TOF_BETATPREPPAIE.InitScreenObject;
begin

  if chkApercu <> nil then ChkApercu.Checked   := True;
  if ChkReduire <> nil then ChkReduire.Checked  := False;
  if ChkCouleur <> nil then ChkCouleur.Checked  := False;
  if ChkDefinitif <> nil then ChkDefinitif.Checked:= False;

  if Fetat <> nil then
  begin
    FEtat.Items.Clear;
    FEtat.Values.Clear;
  end;

end;

Procedure TOF_BETATPREPPAIE.ControleChamp(Champ : String;Valeur : String);
Begin

end;

procedure TOF_BETATPREPPAIE.ChargeEtat;
begin

  TheType   := 'E';
  TheNature := 'EPP';
  TheModele := '';
  TheTitre  := 'Etat Préparatoire de Paie';

  OptionEdition := TOptionEdition.Create(TheType, TheNature, TheModele, TheTitre, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, TPageControl(Ecran.FindComponent('Pages')), fEtat);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat,Idef);

  if fetat.itemindex  >= 0 then TheModele := FETAT.values[fetat.itemindex];

end;

procedure TOF_BETATPREPPAIE.CreateTobs;
Begin

  TobConso   := TOB.Create('LES CONSOS', nil, -1);

end;

Procedure TOF_BETATPREPPAIE.OnChangeEtat(Sender : Tobject);
Begin

  if THValComboBOx(sender).name = 'FETAT' then
  begin
   OptionEdition.Modele := FETAT.values[fetat.itemindex];
   TheModele  := OptionEdition.Modele;
  end;

end;

procedure TOF_BETATPREPPAIE.OnClickApercu(Sender : Tobject);
begin
  OptionEdition.Apercu  := ChkApercu.checked;
end;

procedure TOF_BETATPREPPAIE.OnClickReduire(Sender : Tobject);
begin
  OptionEdition.DeuxPages  := ChkReduire.checked;
end;

procedure TOF_BETATPREPPAIE.OnClickCouleur(Sender : Tobject);
begin
  OptionEdition.Couleur  := ChkCouleur.checked;
end;

procedure TOF_BETATPREPPAIE.OnClickDefinitif(Sender : Tobject);
begin
  //edition définitive des rubriques de paie...
end;

procedure TOF_BETATPREPPAIE.ReleaseTob;
Begin
  //
  FreeAndNil(TobCONSO);
  //
end;

//appel de l'écran de modificatuion/création d'état
procedure TOF_BETATPREPPAIE.ParamEtat(Sender : TOBJect);
begin

  OptionEdition.Appel_Generateur;

end;

procedure TOF_BETATPREPPAIE.OnClickValide (Sender : Tobject);
var StSQL   : String;
    ResArg  : String;
    QQ      : TQuery;
    TobSemaine : TOB;
    TOBEdition : TOB;
    Ind     : Integer;
begin

  CreateTobs;

  TobEditions   := TOB.create('LES EDITIONS', nil, -1);

  //chargement des ressources associées à une consommation pour la période en cours
  StSql := 'SELECT ARS_TYPERESSOURCE, ARS_RESSOURCE, ARS_LIBELLE, ARS_LIBELLE2, CONSOMMATIONS.*, ';
  StSql := StSql + 'GA_TYPEARTICLE, GA_LIBELLE, GA_FAMILLENIV1, GA_FAMILLENIV2, GA_FAMILLENIV3, AFF_TIERS ';
  stSql := StSql + 'FROM CONSOMMATIONS ';
  stSql := StSql + 'LEFT JOIN RESSOURCE ON BCO_RESSOURCE=ARS_RESSOURCE ';
  stSql := StSql + 'LEFT JOIN ARTICLE ON BCO_ARTICLE=GA_ARTICLE ';
  stSql := StSql + 'LEFT JOIN AFFAIRE ON BCO_AFFAIRE=AFF_AFFAIRE ';
  stSql := StSql + 'WHERE BCO_PAIENUMFIC = 0 AND ARS_TYPERESSOURCE="SAL" AND BCO_NATUREMOUV IN ("MO", "FRS") AND';

  //vérification si les codes ressources sont pas renseignés
  ResARg := '';
  If (ressource.Text <> '') And (Ressource1.Text <> '') then
  begin
    if Ressource.text = Ressource1.Text then
    begin
      ResArg := ' BCO_RESSOURCE ="' + Ressource.text + '" AND ';
    end
    Else
    begin
      ResArg := ' BCO_RESSOURCE >="' + Ressource.Text + '" AND BCO_RESSOURCE <= "' + Ressource1.text + '" AND ';
    end;
    StSQL := StSQL + ResArg;
  end;

  //chargement des dates de début et de fin pour complément Requete
  ResArg := ' BCO_DATEMOUV >= "' + USDATETIME(StrToDateTime(DateDeb.text)) + '" AND BCO_DATEMOUV <= "' + USDATETIME(StrToDateTime(Datefin.text)) + '" ';
  StSQL := StSQL + ResArg + ' ORDER BY BCO_RESSOURCE, BCO_ARTICLE, BCO_SEMAINE';

  QQ := OpenSQL(StSQL, false);

  if not QQ.eof then
  begin
    TobCONSO.LoadDetailDB('CONSOMMATIONS ', '','',QQ, false);
    //chargement de la tob globale d'édition
    TobEdition := TOB.Create('LEDITIONS', TobEditions, -1);
    //chargement de l'entête de la Tob d'édition
    TobEdition.AddChampSupValeur('DATEDEB', DateDeb.Text);
    TobEdition.AddChampSupValeur('DATEFIN', DateFin.Text);

    if ChkDefinitif.checked then
      TobEdition.AddChampSupValeur('DEFINITIF', '1')
    Else
      TobEdition.AddChampSupValeur('DEFINITIF', '0');;

    TobEdition.AddChampSupValeur('SEMAINE1', 0);
    TobEdition.AddChampSupValeur('SEMAINE2', 0);
    TobEdition.AddChampSupValeur('SEMAINE3', 0);
    TobEdition.AddChampSupValeur('SEMAINE4', 0);
    TobEdition.AddChampSupValeur('SEMAINE5', 0);

    //chargement des semaines à éditer
    TobSemaine := TOB.Create('LES SEMAINES', nil, -1);
    StSQL := 'SELECT DISTINCT BCO_SEMAINE FROM CONSOMMATIONS WHERE ' + ResArg + ' ORDER BY BCO_SEMAINE';
    QQ := OpenSQL(StSQL, false);
    TobSemaine.LoadDetailDB('LA SEMAINE', '','',QQ, false);
    For ind := 0 TO TobSemaine.detail.count -1 do
    begin
      if      Ind = 0 then TobEdition.PutValue('SEMAINE1', TobSemaine.Detail[Ind].GetString('BCO_SEMAINE'))
      else if Ind = 1 then TobEdition.PutValue('SEMAINE2', TobSemaine.Detail[Ind].GetString('BCO_SEMAINE'))
      Else if Ind = 2 then TobEdition.PutValue('SEMAINE3', TobSemaine.Detail[Ind].GetString('BCO_SEMAINE'))
      Else if Ind = 3 then TobEdition.PutValue('SEMAINE4', TobSemaine.Detail[Ind].GetString('BCO_SEMAINE'))
      Else if Ind = 4 then TobEdition.PutValue('SEMAINE5', TobSemaine.Detail[Ind].GetString('BCO_SEMAINE'))
      Else if Ind > 4 Then TobEdition.PutValue('SEMAINE5', 'CUMULEES')
      Else TobEdition.PutValue('SEMAINE5', 'ERREURS SEMAINES');
    end;
    //
    ChargementLigneEdition(TobEdition);
    //
    EditeTout;
    //
    FreeAndNil(TobEditions);
  end;

  Ferme(QQ);

  TToolbarButton97(ecran.FindComponent('BValider')).Down := False;

  if ChkDefinitif.checked then
  begin
    StsQL :=  'UPDATE CONSOMMATIONS SET BCO_PAIENUMFIC=1 FROM CONSOMMATIONS ';
    stSql := StSql + 'LEFT JOIN RESSOURCE ON BCO_RESSOURCE=ARS_RESSOURCE ';
    stSql := StSql + 'LEFT JOIN ARTICLE ON BCO_ARTICLE=GA_ARTICLE ';
    stSql := StSql + 'LEFT JOIN AFFAIRE ON BCO_AFFAIRE=AFF_AFFAIRE ';
    stSql := StSql + 'WHERE BCO_PAIENUMFIC = 0 AND ARS_TYPERESSOURCE="SAL" AND ';
    //
    StSQL := StSQL + ResArg;
    executesql(StSQL);
  end;

  ReleaseTob;

end;

procedure TOF_BETATPREPPAIE.ChargementLigneEdition(TobEdition : TOB);
Var Ind     : Integer;
    TOBLIG  : TOB;
    TOBTemp : TOB;
    //
    DateMouv: TDateTime;
    //
    iYear   : Word;
    iMonth  : Word;
    iDay    : Word;
    //
    CumulQte: Double;
    CumulH  : Double;
    CumulAC : Double;
    CumulPR : Double;
    CumulHT : Double;
    //
    xxx     : String;
begin

    for Ind := 0 to TobConso.Detail.count - 1 do
    begin
      TobLig   := TobCONSO.Detail[Ind];
      DateMouv := TobLig.GetDateTime('BCO_DATEMOUV');
      Decodedate(DateMouv, iyear, imonth, iday);
      //contrôle si cette ressource est déjà sur une semaine dans la table
      //TobTemp := TobEdition.FindFirst(['BCO_RESSOURCE','SEMAINE','ARTICLE'], [TobLig.GetString('BCO_RESSOURCE'), TobLig.GetString('BCO_SEMAINE'), TobLig.GetString('BCO_ARTICLE')], false);
      //TobTemp := TobEdition.FindFirst(['BCO_RESSOURCE','SEMAINE'], [TobLig.GetString('BCO_RESSOURCE'), TobLig.GetString('BCO_SEMAINE')], false);
      xxx := TOBLig.GetString('BCO_TYPEHEURE');
      TobTemp := TobEdition.FindFirst(['BCO_RESSOURCE','ARTICLE', 'TYPEHEURE'], [TobLig.GetString('BCO_RESSOURCE'),TobLig.GetString('BCO_ARTICLE'), TobLig.GetString('BCO_TYPEHEURE')], false);
      if TobTemp <> nil then
      begin
        //cumul sur la ligne des valeurs et montants...
        If TobLig.GetString('BCO_NATUREMOUV') = 'MO' then
        begin
          if TOBLig.GetString('BCO_SEMAINE')= TobEdition.GetString('SEMAINE1') THEN
          begin
            CumulH  := TobTemp.GetDouble('QTESEM1') + TobLig.GetDouble('BCO_QUANTITE');
            TobTemp.SetDouble('QTESEM1', CumulH);
          end
          else if TOBLig.GetString('BCO_SEMAINE')= TobEdition.GetString('SEMAINE2') THEN
          begin
            CumulH  := TobTemp.GetDouble('QTESEM2') + TobLig.GetDouble('BCO_QUANTITE');
            TobTemp.SetDouble('QTESEM2', CumulH);
          end
          else if TOBLig.GetString('BCO_SEMAINE') = TobEdition.GetString('SEMAINE3') THEN
          begin
            CumulH  := TobTemp.GetDouble('QTESEM3') + TobLig.GetDouble('BCO_QUANTITE');
            TobTemp.SetDouble('QTESEM3', CumulH);
          end
          else if TOBLig.GetString('BCO_SEMAINE') = TobEdition.GetString('SEMAINE4') THEN
          begin
            CumulH  := TobTemp.GetDouble('QTESEM4') + TobLig.GetDouble('BCO_QUANTITE');
            TobTemp.SetDouble('QTESEM4', CumulH);
          end
          else
          begin
            CumulH  := TobTemp.GetDouble('QTESEM5') + TobLig.GetDouble('BCO_QUANTITE');
            TobTemp.PutValue('QTESEM5', CumulH);
          end;
          //
          CumulAC := TobTemp.GetDouble('MONTANTACH')+ TobLig.GetDouble('BCO_MONTANTACH');
          CumulPR := TobTemp.GetDouble('MONTANTPR') + TobLig.GetDouble('BCO_MONTANTPR');
          CumulHT := TobTemp.GetDouble('MONTANTHT') + TobLig.GetDouble('BCO_MONTANTHT');
          //
          TobTemp.SetDouble('MONTANTACH', CumulAC);
          TobTemp.SetDouble('MONTANTPR',  CumulPR);
          TobTemp.SetDouble('MONTANTHT',  CumulHT);
        end
        else
        begin
          //Création de la ligne ressource
          //CreateLigneEdition(TOBEdition, TobLig, iMonth);                    //
          CumulQte:= TobTemp.GetDouble('QUANTITE')  + TobLig.GetDouble('BCO_QUANTITE');
          CumulAC := TobTemp.GetDouble('MONTANTACH')+ TobLig.GetDouble('BCO_MONTANTACH');
          CumulPR := TobTemp.GetDouble('MONTANTPR') + TobLig.GetDouble('BCO_MONTANTPR');
          CumulHT := TobTemp.GetDouble('MONTANTHT') + TobLig.GetDouble('BCO_MONTANTHT');
          //
          TobTemp.SetDouble('QUANTITE',   CumulQte);
          TobTemp.SetDouble('MONTANTACH', CumulAC);
          TobTemp.SetDouble('MONTANTPR',  CumulPR);
          TobTemp.SetDouble('MONTANTHT',  CumulHT);
          //
        end;
      end
      else
      begin
        CreateLigneEdition(TOBEdition, TobLig, iMonth);
      end;
    end;

    TOBEdition.Detail.Sort('BCO_RESSOURCE;TYPELIG;ARTICLE;TYPEHEURE');

end;

procedure TOF_BETATPREPPAIE.CreateLigneEdition(TOBEdition,Toblig : TOB; iMonth : word);
Var TobAdd : TOB;
begin

  TobAdd := TOB.Create('Une Ligne', TobEdition, -1);

  TobAdd.AddChampSupValeur('BCO_RESSOURCE', Toblig.GetString('BCO_RESSOURCE'));
  TobAdd.AddChampSupValeur('NOM',     Toblig.GetString('ARS_LIBELLE'));
  TobAdd.AddChampSupValeur('PRENOM',  Toblig.GetString('ARS_LIBELLE2'));
  //
  TOBAdd.AddChampSupValeur('MOIS', iMonth);
  //TobAdd.AddChampSupValeur('SEMAINE', Toblig.GetString('BCO_SEMAINE'));
  //
  TOBAdd.AddChampSupValeur('TYPELIG', '');
  TOBAdd.AddChampSupValeur('TYPEMOUV','');
  TobAdd.AddchampSupValeur('TYPEHEURE', '');
  //
  TOBAdd.AddChampSupValeur('ARTICLE',     TobLig.GetString('BCO_ARTICLE'));
  TOBAdd.AddchampSupValeur('DESIGNATION', TobLig.GetString('GA_LIBELLE'));
  //
  TOBAdd.AddChampSupValeur('QTESEM1', 0);
  TOBAdd.AddChampSupValeur('QTESEM2', 0);
  TOBAdd.AddChampSupValeur('QTESEM3', 0);
  TOBAdd.AddChampSupValeur('QTESEM4', 0);
  TOBAdd.AddChampSupValeur('QTESEM5', 0);

  If TobLig.GetString('BCO_NATUREMOUV') = 'MO' then
  Begin
    TOBAdd.SetString('TYPELIG',   'H');
    TOBAdd.SetString('TYPEHEURE',  TOBLig.GetString('BCO_TYPEHEURE'));
    TOBAdd.AddChampSupValeur('QUANTITE',0);
    if TOBAdd.GetString('TYPEHEURE') <> '' then
      TOBAdd.SetString('TYPEMOUV',  TOBLig.GetString('BCO_NATUREMOUV') + '-' + TOBAdd.GetString('TYPEHEURE'))
    else
      TOBAdd.SetString('TYPEMOUV',  TOBLig.GetString('BCO_NATUREMOUV'));
    //
    if TobEdition.GetString('SEMAINE1') = TOBLig.GetString('BCO_SEMAINE') THEN
      TOBAdd.SetDouble('QTESEM1',    Toblig.GetDouble('BCO_QUANTITE'))
    Else if TobEdition.GetString('SEMAINE2') = TOBLig.GetString('BCO_SEMAINE') THEN
      TOBAdd.SetDouble('QTESEM2',    Toblig.GetDouble('BCO_QUANTITE'))
    Else if TobEdition.GetString('SEMAINE3') = TOBLig.GetString('BCO_SEMAINE') THEN
      TOBAdd.SetDouble('QTESEM3',    Toblig.GetDouble('BCO_QUANTITE'))
    Else if TobEdition.GetString('SEMAINE4') = TOBLig.GetString('BCO_SEMAINE') THEN
      TOBAdd.SetDouble('QTESEM4',    Toblig.GetDouble('BCO_QUANTITE'))
    Else if TobEdition.GetString('SEMAINE5') = TOBLig.GetString('BCO_SEMAINE') THEN
      TOBAdd.SetDouble('QTESEM5',    Toblig.GetDouble('BCO_QUANTITE'))
    Else
      TOBAdd.SetDouble('QTESEM5',    Toblig.GetDouble('BCO_QUANTITE'));
  end
  else
  Begin
    TOBAdd.SetString('TYPELIG',   'Z');
    TOBAdd.SetString('TYPEMOUV',  TOBLig.GetString('BCO_NATUREMOUV') + '/' + TOBLig.GetString('BCO_CODEARTICLE'));
    TOBAdd.AddChampSupValeur('QUANTITE',  Toblig.GetDouble('BCO_QUANTITE'));
  end;
  //
  TOBAdd.AddChampSupValeur('PU',          Toblig.GetDouble('BCO_PUHT'));
  TOBAdd.AddChampSupValeur('DPA',         Toblig.GetDouble('BCO_DPA'));
  TOBAdd.AddChampSupValeur('DPR',         Toblig.GetDouble('BCO_DPR'));
  TOBAdd.AddChampSupValeur('MONTANTACH',  Toblig.GetDouble('BCO_MONTANTACH'));
  TOBAdd.AddChampSupValeur('MONTANTPR',   Toblig.GetDouble('BCO_MONTANTPR'));
  TOBAdd.AddChampSupValeur('MONTANTHT',   Toblig.GetDouble('BCO_MONTANTHT'));

end;

procedure TOF_BETATPREPPAIE.EditeTout;
var OptEdt : TOptionEdition;
begin

  OptEdt := TOptionEdition.create;
  if TOBEditions.detail.count = 0 then exit;

  OptEdt.Tip     := TheType;
  OptEdt.Nature  := TheNature;
  OptEdt.Modele  := TheModele;
  OptEdt.Titre   := TheTitre;
  //
  OptEdt.Apercu     := ChkApercu.Checked;
  OptEdt.DeuxPages  := ChkReduire.Checked;
  OptEdt.Spages     := TPageControl(Ecran.FindComponent('Pages'));
  //
  if OptEdt.LanceImpression('', TobEditions) < 0 then V_PGI.IoError:=oeUnknown;

  OptEdt.free;

end;

Initialization
  registerclasses ( [ TOF_BETATPREPPAIE ] ) ;
end.

