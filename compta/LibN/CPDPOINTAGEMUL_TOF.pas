{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/09/2003
Modifié le ... : 26/09/2006
Description .. : Source TOF de la FICHE : CPDPOINTAGEMUL ()
Suite ........ : 
Suite ........ : 26/09/06 : JP : mise en place du multi sociétés Treso 
Mots clefs ... : TOF;CPDPOINTAGEMUL
*****************************************************************}

Unit CPDPOINTAGEMUL_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls, 
    Classes,
{$IFDEF EAGLCLIENT}
    Maineagl,      // AGLLanceFiche EAGL
    eMul,          // TFMUL(ECRAN).BChercheClick( nil )
{$ELSE}
    Fe_main,       // AGLLanceFiche AGL
    Mul,           // TFMUL(ECRAN).BChercheClick( nil )
    DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    Hdb,
{$ENDIF}
{$IFDEF VER150}
    Variants,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    Ent1,            // VH^.
    Htb97,           // TToolBarButton97
    uTob,            // TOB
    Windows,         // VK_
    Menus,           // TPopUpMenu
    LookUp,          // LookUpList
    CPDPOINTAGE_TOF  // CPLanceFiche_DPointage
    ;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPDPOINTAGEMUL = Class (TOF)
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;

    public
{$IFDEF EAGLCLIENT}
        FListe : THGrid;
{$ELSE}
        FListe : THDBGrid;
{$ENDIF}

        TEE_GENERAL     : THLabel;
        EE_GENERAL      : THEdit;
        EE_DATEPOINTAGE : THEdit;
        EE_REFPOINTAGE  : THEdit;

        Q           : TQuery;

        BOuvrir     : TToolBarButton97;
        BCherche    : TToolBarButton97;
        BRechercher : TToolBarButton97;

        POPF11   : TPopUpMenu;
        {$IFDEF TRESO}
        NomBase  : string;
        BqCode   : string;
        {$ENDIF TRESO}
    private
        FStArgument : string;

        procedure OnDblClickFListe(Sender : TObject);
        {$IFDEF TRESO}
        procedure BqCodeOnElipsisClick(Sender : TObject);
        procedure BqCodeOnChange(Sender : TObject);
        {$ENDIF TRESO}
        procedure OnElipsisClickEE_General(Sender : TObject);
        procedure OnElipsisClickEE_DatePointage(Sender : TObject);
        procedure OnElipsisClickEE_RefPointage(Sender : TObject);

        procedure OnKeyDownEcran(Sender : TObject ; var Key : Word ; Shift : TShiftState);
end ;

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
procedure CPLanceFiche_DPointageMul;

//================================================================================
// Implementation
//================================================================================
Implementation

{$IFDEF TRESO}
uses
  Constantes, Commun, UTilPgi;
{$ENDIF TRESO}

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_DPointageMul;
begin
    AGLLanceFiche('CP','CPDPOINTAGEMUL','','','');
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnLoad;
var
    lStArg : string;
begin
    Inherited;

    if (FStArgument <> '') then
    begin
        lStArg := ReadTokenSt(FStArgument);
        if (lStarg <> '') then
        begin
            EE_General.Text := ReadTokenSt(lStArg);
            EE_General.Enabled := False;
            Ecran.WindowState := wsMaximized;
        end;
    end;

    if (not VH^.PointageJal) then
    begin
        Ecran.Caption        := 'Annulation du pointage sur compte';
        TEE_GENERAL.Caption  := 'Compte général';
        EE_GENERAL.MaxLength := VH^.CPta[fbGene].Lg;
    end
    else
    begin
        Ecran.Caption        := 'Annulation du pointage sur journal';
        TEE_GENERAL.Caption  := 'Code journal';
        EE_GENERAL.MaxLength := 3;
    end;
    Ecran.Caption:=TraduireMemoire(Ecran.Caption) ;
    TEE_GENERAL.Caption  :=TraduireMemoire(TEE_GENERAL.Caption) ;
    UpDateCaption(Ecran);
  {$IFDEF TRESO}
  {JP 02/11/06 : En tréso Multi sociétés, on ne "tape" pas nécessairement dans la base courante}
  if IsTresoMultiSoc then begin
    SetControlText('XX_FROM', GetTableDossier(NomBase, 'EEXBQ'));
  end;
  {$ENDIF TRESO}

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnArgument(S : String);
begin
    Inherited;
    FStArgument := S;
    Ecran.HelpContext := 7607000 ;
    TEE_GENERAL     := THLabel(GetControl('TEE_GENERAL',True));
    EE_GENERAL      := THEdit(GetControl('EE_GENERAL',True));
    EE_DATEPOINTAGE := THEdit(GetControl('EE_DATEPOINTAGE',True));
    EE_REFPOINTAGE  := THEdit(GetControl('EE_REFPOINTAGE',True));

    EE_GENERAL.OnElipsisClick      := OnElipsisClickEE_GENERAL;
    EE_DATEPOINTAGE.OnElipsisClick := OnElipsisClickEE_DATEPOINTAGE;
    EE_REFPOINTAGE.OnElipsisClick  := OnElipsisClickEE_REFPOINTAGE;
    {$IFDEF TRESO}
    if IsTresoMultiSoc then begin
      TFMul(Ecran).SetDBListe('TREEXBQ');
      SetControlVisible('EE_GENERAL', False);
      SetControlVisible('BQCODE', True);
      THEdit(GetControl('BQCODE',True)).OnElipsisClick := BqCodeOnElipsisClick;
      THEdit(GetControl('BQCODE',True)).OnChange := BqCodeOnChange;
    end
    else
      NomBase := V_PGI.SchemaName;
    {$ENDIF TRESO}

{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FLISTE',True));
{$ELSE}
    FListe := THDBGrid(GetControl('FLISTE',True));
{$ENDIF}
    FListe.OnDblClick := OnDblClickFListe;

    BCherche    := TToolBarButton97(GetControl('BCHERCHE',True));
    BOuvrir     := TToolBarButton97(GetControl('BOUVRIR',True));
    BRechercher := TToolBarButton97(GetControl('BRECHERCHER',True));

    POPF11 := TPopUpMenu(GetControl('POPF11',True));
    POPF11.Items[0].OnClick := OnDblClickFListe;

    AddMenuPop(PopF11,'','');

    Ecran.OnKeyDown := OnKeyDownEcran;

    Q := TQuery(GetControl('Q',True));
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnDblClickFListe(Sender : TObject);
begin
    if (GetField('EE_GENERAL') = null) then Exit;
    {$IFDEF TRESO}
    CPLanceFiche_DPointage(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + NomBase + ';');
    {$ELSE}
    CPLanceFiche_DPointage(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE'));
    {$ENDIF TRESO}
    TFMUL(ECRAN).BChercheClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnElipsisClickEE_General(Sender : TObject);
begin
  {$IFDEF TRESO}
  if (not VH^.PointageJal) then begin
    if not IsTresoMultiSoc then
      LookupList(THEdit(Sender), 'Compte bancaire', 'BANQUECP', 'BQ_GENERAL', 'BQ_LIBELLE', FiltreBanqueCp('', tcb_Bancaire, ''), 'BQ_CODE', True, 0);
  end
  {$ELSE}
  if (not VH^.PointageJal) then LookUpList(THEdit(Sender),'Compte général','GENERAUX','G_GENERAL','G_LIBELLE','G_POINTABLE="X"','G_GENERAL',True,0)
  {$ENDIF TRESO}
  else LookUpList(THEdit(Sender),'Journal','JOURNAL','J_JOURNAL','J_LIBELLE','','J_JOURNAL',True,0,'SELECT J_JOURNAL,J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X" ORDER BY J_JOURNAL');
end;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOF_CPDPOINTAGEMUL.BqCodeOnElipsisClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookupList(THEdit(Sender), 'Comptes bancaires', 'BANQUECP', 'BQ_CODE', 'BQ_LIBELLE', FiltreBanqueCp('', tcb_Bancaire, ''), 'BQ_CODE', True, 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPDPOINTAGEMUL.BqCodeOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  BqCode := GetControlText('BQCODE');
  Q := OpenSQL('SELECT BQ_GENERAL, DOS_NOMBASE FROM BANQUECP ' +
               'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
               'WHERE BQ_CODE = "' + BQCode + '"', True);
  if not Q.EOF then begin
    NomBase := Q.FindField('DOS_NOMBASE').AsString;
    SetControlText('EE_GENERAL', Q.FindField('BQ_GENERAL').AsString);
  end;
  Ferme(Q);
end;
{$ENDIF TRESO}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnElipsisClickEE_DatePointage(Sender : TObject);
begin
  {$IFDEF TRESO}
  LookUpList(THEdit(Sender),'Date de pointage', GetTableDossier(NomBase, 'EEXBQ'),'EE_DATEPOINTAGE','EE_REFPOINTAGE','EE_GENERAL="' + EE_General.Text + '"','EE_DATEPOINTAGE DESC',True,0);
  {£ELSE}
  LookUpList(THEdit(Sender),'Date de pointage', 'EEXBQ', 'EE_DATEPOINTAGE','EE_REFPOINTAGE','EE_GENERAL="' + EE_General.Text + '"','EE_DATEPOINTAGE DESC',True,0);
  {$ENDIF TRESO}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnElipsisClickEE_RefPointage(Sender : TObject);
begin
  {$IFDEF TRESO}
  LookUpList(THEdit(Sender),'Référence de pointage',GetTableDossier(NomBase, 'EEXBQ'),'EE_REFPOINTAGE','EE_DATEPOINTAGE','EE_GENERAL="' + EE_General.Text + '"','EE_DATEPOINTAGE DESC',True,0);
  {$ELSE}
  LookUpList(THEdit(Sender),'Référence de pointage', 'EEXBQ','EE_REFPOINTAGE','EE_DATEPOINTAGE','EE_GENERAL="' + EE_General.Text + '"','EE_DATEPOINTAGE DESC',True,0);
  {$ENDIF TRESO}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGEMUL.OnKeyDownEcran(Sender : TObject ; var Key : Word ; Shift : TShiftState);
begin
    case key of
        //Ctrl + H
        70 : if (Shift = [ssCtrl]) then BRechercher.Click;
       (* 26/09/06 : Cela me parait inutile 
        VK_F5 :
            begin
                if (EE_General.Focused) then EE_General.ElipsisClick(Self)
                else if (EE_DatePointage.Focused) then EE_DatePointage.ElipsisClick(Self)
                else if (EE_RefPointage.Focused) then EE_RefPointage.ElipsisClick(Self)
                else if (FListe.Focused) then BOuvrir.Click;
                {$IFDEF TRESO}
                Key := 0;
                {Nécessaire, sinon le lookup est appelè deux fois}
                //if IsTresoMultiSoc then
                {$ENDIF TRESO}
            end;

        VK_F9 :
            begin
                BCherche.Click;
                if (FListe.CanFocus) then FListe.SetFocus;
            end;
       *)
        VK_F10 : BOuvrir.Click;
        
        VK_F11 : POPF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);
        
        VK_F12 : if (FListe.Focused) then TPageControl(GetControl('PAGES', False)).SetFocus
                 else FListe.SetFocus;
    end;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPDPOINTAGEMUL]); 
end.
