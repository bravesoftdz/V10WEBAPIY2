{***********UNITE*************************************************
Auteur  ...... : MD
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYMULUSERCONF
Mots clefs ... : TOF;USERCONF
*****************************************************************}
Unit galTofUserConf ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul, MaineAGL, Utob,
{$ELSE}
     db,
     dbtables, Mul, FE_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97, AGLInit, HDB, HQry ;

Type
  TOF_USERCONF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Lst          : THDBGrid;            // Liste du mul
    Qry          : THQuery;             // Query du mul
    ModeModif    : Boolean ;            // True si mode modif des conf (False si consult)
    mvGroupeConf : THMultiValComboBox ; // groupes de confidentialité
    procedure ResetTimer(Sender: TObject);
    procedure CHKSANSGROUPE_OnClick(Sender: TObject);
    procedure MVGROUPECONF_OnChange(Sender: TObject);
    procedure BOUVRIR_OnClick(Sender: TObject);
    procedure BDELETE_OnClick(Sender: TObject);
    procedure BMODIFGROUPES_OnClick(Sender: TObject);
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;


///////////// IMPLEMENTATION //////////////
Implementation

uses
{$IFDEF EAGLCLIENT}
  eTablette,
{$ELSE}
  Tablette,
{$ENDIF}
  galAffectUserConf, UtilMulTraitmt,windows;

procedure TOF_USERCONF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_USERCONF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_USERCONF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_USERCONF.OnLoad;
begin
  inherited;
end;

procedure TOF_USERCONF.OnArgument (S : String ) ;
var tmp : String;
begin
  Inherited ;
  // La fiche YYMULUSERCONF peut recevoir 'MODEMODIF' en paramètre
  ModeModif := False;
  tmp := ReadTokenSt(S);
  if Copy(tmp, 1, 7)='ACTION=' then tmp := ReadTokenSt(S);
  if tmp='MODEMODIF' then ModeModif := True;

  // récup des objets
  Lst := TFMul(Ecran).FListe;
  Qry := TFMul(Ecran).Q;

  // Evènements
  Ecran.OnKeyDown := Form_OnKeyDown;
  mvGroupeConf := THMultiValComboBox(GetControl('MVGROUPECONF'));
  mvGroupeConf.OnChange := MVGROUPECONF_OnChange;
  TCheckBox(GetControl('CHKSANSGROUPE')).OnClick := CHKSANSGROUPE_OnClick;
  TToolBarButton97(GetControl('BDELETE')).OnClick := BDELETE_OnClick;
  TToolBarButton97(GetControl('BMODIFGROUPES')).OnClick := BMODIFGROUPES_OnClick;
  // pour les zones non liées à un nom de champ, gestion des modes
  Case TFMul(Ecran).AutoSearch of
    asChange : ; // voir onclick de la case
    asExit   :   // voir proc. InitAutoSearch du mul
      begin
      TCheckBox(GetControl('CHKSANSGROUPE')).onExit := TFMul(Ecran).SearchTimerTimer ;
      mvGroupeConf.OnExit := TFMul(Ecran).SearchTimerTimer ;
      end;
    asTimer  : ; // voir onclick de la case
  end;

  // en modemodif, on doit avoir une liste d'utilisateurs
  // (et non un produit cartésien dû aux croisements avec userconf)
  if ModeModif then
    begin
    Ecran.Caption := 'Affectation des utilisateurs à des groupes';
    UpdateCaption(Ecran);
    // donc on passe sur une liste affichant des colonnes de UTILISAT uniquement
    TFMul(Ecran).Q.Manuel := True;
    // TFMul(Ecran).Q.Liste := 'YYMULUSERCONF2';
    TFMul(Ecran).SetDBListe('YYMULUSERCONF2');
    TFMul(Ecran).Q.Manuel := False;
    // et on interdit l'accès aux autres champs de la vue
    SetControlVisible('BParamListe', False);
    // et on ne fait pas de suppression ... d'utilisateurs !
    SetControlVisible('BDELETE', False);
    TToolBarButton97(GetControl('BOUVRIR')).OnClick := BOUVRIR_OnClick;
    // mais on peut modifier les groupes
    SetControlVisible('BMODIFGROUPES', True);
    end
  else
    begin
    SetControlVisible('BOUVRIR', False);
    SetControlVisible('BINSERT', True);
    TToolBarButton97(GetControl('BINSERT')).OnClick := BOUVRIR_OnClick;
    end;
end ;

procedure TOF_USERCONF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_USERCONF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_USERCONF.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_USERCONF.ResetTimer(Sender: TObject);
begin
 // idem proc dans Mul
 TFMul(Ecran).SearchTimer.Enabled := False ;
 TFMul(Ecran).SearchTimer.Enabled := True ;
end;


procedure TOF_USERCONF.CHKSANSGROUPE_OnClick(Sender: TObject);
var XX_Where: String;
begin
  XX_Where := '';
  if TCheckBox(GetControl('CHKSANSGROUPE')).checked then
    begin
    SetControlText('MVGROUPECONF', '');
    SetControlEnabled('MVGROUPECONF', False);
    // on a le champ dans la liste YYMULUSERCONF basée sur la vue du même nom
    if Not ModeModif then
      XX_Where := ' and (UCO_GROUPECONF is null)'
    // on n'a que la table UTILISAT dans la liste YYMULUSERCONF2
    else
      XX_Where := ' and US_UTILISATEUR not in (select UCO_USER from USERCONF)';
    end
  else
    SetControlEnabled('MVGROUPECONF', True);

  // maj du critère global
  SetControlText('XX_WHERE', XX_Where);

  // voir proc. InitAutoSearch du mul
  Case TFMul(Ecran).AutoSearch of
    asChange : TFMul(Ecran).SearchTimerTimer(Sender) ;
    asTimer  : ResetTimer(Sender) ;
  end;
end;


procedure TOF_USERCONF.MVGROUPECONF_OnChange(Sender: TObject);
var XX_Where, sGroupeConfText, tmp, liaison: String;
begin
  XX_Where := '';
  sGroupeConfText := mvGroupeConf.Text;
  if sGroupeConfText<>'' then
    begin
    tmp := ReadTokenSt(sGroupeConfText);
    if (tmp<>'<<Tous>>') then // ou mvGroupeConf.Tous = True ?
      BEGIN
      // on a le champ dans la liste YYMULUSERCONF basée sur la vue du même nom
      if ModeModif then
        // on n'a que la table UTILISAT dans la liste YYMULUSERCONF2
        begin
        XX_Where := ' and US_UTILISATEUR in (select UCO_USER from USERCONF ';
        liaison := ' where ';
        end
      else
        liaison := ' and (';
      while tmp<>'' do
        begin
        // tmp contient le Code du groupe de conf.
        XX_Where := XX_Where + liaison + 'UCO_GROUPECONF="'+tmp+'"';
        liaison := ' or ';
        tmp := ReadTokenSt(sGroupeConfText);
        end;
      XX_Where := XX_Where + ')';
      END;
    end;

  // maj du critère global
  SetControlText('XX_WHERE', XX_Where);

  // voir proc. InitAutoSearch du mul
  Case TFMul(Ecran).AutoSearch of
    asChange : TFMul(Ecran).SearchTimerTimer(Sender) ;
    asTimer  : ResetTimer(Sender) ;
  end;
end;


procedure TOF_USERCONF.BOUVRIR_OnClick(Sender: TObject);
begin
  // Depuis la consultation, la mouette verte ouvre la même fiche en mode modif
  if Not ModeModif then
    begin
    AGLLanceFiche('YY','YYMULUSERCONF','','','MODEMODIF');
    // suite aux maj : on doit faire un refreshdb
    AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
    exit;
    end;

  // En mode modif, ouvre la fiche d'affectation des groupes
  if (Lst.NbSelected=0) and (not Lst.AllSelected) then
    begin
    PGIInfo('Aucun utilisateur sélectionné.', TitreHalley);
    exit;
    end;

  FicheAffectUserConf(Ecran);

  // suite aux maj : on doit faire un refreshdb
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_USERCONF.BDELETE_OnClick(Sender: TObject);
var codgrp, coduser: String;
    i : Integer;
    procedure SupprimeLenregUserConf;
    begin
      codgrp := Qry.FindField('UCO_GROUPECONF').AsString;
      coduser := Qry.FindField('US_UTILISATEUR').AsString;
      ExecuteSQL('delete from USERCONF where UCO_GROUPECONF="'+codgrp
                +'" and UCO_USER="'+coduser+'"');
    end;
begin
  if (Lst.NbSelected=0) and (Not Lst.AllSelected) then
  begin
    PGIInfo('Aucun enregistrement sélectionné.', TitreHalley);
    exit;
  end;

  if PGIAsk('Vous allez supprimer définitivement les affectations sélectionnées.'+#13#10
           +' Confirmez vous l''opération ?', TitreHalley)<>mrYes then exit;

  if Lst.AllSelected then
    BEGIN
{$IFDEF EAGLCLIENT}
    if not TFMul(Ecran).FetchLesTous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      Qry.First;
      while Not Qry.EOF do
        begin
        SupprimeLenregUserConf;
        Qry.Next;
        end;
      end;
    END
  ELSE
    BEGIN
    for i:=0 to Lst.NbSelected-1 do
      begin
      Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Qry.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}
      SupprimeLenregUserConf;
      end;
    END;
  // déselectionne
  FinTraitmtMul(TFMul(Ecran));

  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_USERCONF.BMODIFGROUPES_OnClick(Sender: TObject);
begin
  ParamTable('YYGROUPECONF',taCreat,0,Nil,3); // Groupes de conf
  mvGroupeConf.ReLoad;
end;

procedure TOF_USERCONF.Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 case Key of
  VK_F10 : if (TFMul(Ecran).FListe.Focused) and (GetControl('BDELETE').Visible=False) then
            FicheAffectUserConf(Ecran);

  VK_DELETE : if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) and GetControl('BDELETE').Visible then
               BDELETE_OnClick(Nil);

  else TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 end;
end;

Initialization
  registerclasses ( [ TOF_USERCONF ] ) ;
end.

