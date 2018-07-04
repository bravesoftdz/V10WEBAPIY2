{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... : 07/10/2004
Description .. : Source TOM de la TABLE : DEVISE (DEVISE)
Suite ........ : JP 07/10/04 : Maj des tablettes : cf. RechargeTablettes
Suite ........ : 
Mots clefs ... : TOM;DEVISE
*****************************************************************}
unit DEVISE_TOM;

//================================================================================
// Interface
//================================================================================
interface

uses
    Classes,
    Dialogs,
    stdctrls,
    buttons,
    comctrls,
    UTOM,
    UTOB,
    Hctrls,
    HmsgBox,
    Graphics,
    extctrls,
    Sysutils,
{$ifDEF EAGLCLIENT}
    maineagl,
    eFichList,
{$else}
{$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      FE_Main,
      FichList,
    {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
    db,
    dbctrls,
{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    HDB,
{$endif}
    Ent1,
    HEnt1,
{$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      CPCHANCELL_TOF, {JP20/08/03}
    {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
    SaisUtil
    ;

{$IFDEF EAGLCLIENT}
    type tquery = TOB;
{$ENDIF}

//==================================================
// Definition des Constante
//==================================================
const
     msg0  = '0;Devises;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;';
     msg1  = '1;Devises;Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;';
     msg2  = '2;Devises;Vous devez renseigner un code;W;O;O;O;';
     msg3  = '3;Devises;Vous devez renseigner un libellé;W;O;O;O;';
     msg4  = '4;Devises;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;';
     msg5  = 'L''enregistrement est inaccessible';
     msg6  = '6;Devises;Désirez-vous créer les taux de change correspondant à cette devise?;Q;YNC;Y;C;';
     msg7  = '7;Devises;Vous ne pouvez pas supprimer cette devise: des mouvements lui sont associés.;W;O;O;O;';
     msg8  = '8;Devises;Un taux de zéro bloquera la saisie comptable sur cette devise;E;O;O;O;';
     msg9  = '9;Devises;Aucun taux n''a été saisi. Le taux par défaut sera de 1 dans la saisie sur cette devise;E;O;O;O;';
     msg10 = '10;Devises;Vous ne pouvez pas supprimer la devise pivot !;W;O;O;O;';
     msg11 = '11;Devises;Vous ne pouvez pas modifier le nombre de décimal. Il existe des écritures avec cette devise !;W;O;O;O;';
     msg12 = '12;Devises;Vous ne pouvez pas modifier la quotité. Il existe des écritures avec cette devise !;W;O;O;O;';
     msg13 = '13;Devises;Vous ne pouvez pas supprimer cette devise. Il existe des écritures, des pièces ou des tiers avec cette devise !;W;O;O;O;';
     msg14 = '14;Devises;La parité fixe par rapport à l''Euro doit être positive et de 6 chiffres significatifs maximum !;W;O;O;O;';
     msg15 = '15;Devises;Vous ne pouvez plus modifier la parité. Il existe des écritures avec cette devise !;W;O;O;O;';
     msg16 = '(A saisir sous la forme 1 Euro = xx,xxxxxx';
     msg17 = '17;Devises;Vous ne pouvez pas supprimer la devise pivot ou nationale;W;O;O;O;';

//==================================================
// Definition de Type
//==================================================
type
    CptEtLibelle = record
        compte  : string;
        libelle : string;
    end;

//==================================================
// Definition de class
//==================================================
// TOM_DEVISE
type
    TOM_DEVISE = class(TOM)
        procedure OnNewRecord                ; override ;
        procedure OnDeleteRecord             ; override ;
        procedure OnUpdateRecord             ; override ;
        procedure OnAfterUpdateRecord        ; override ;
        procedure OnLoadRecord               ; override ;
        procedure OnChangeField(F : TField)  ; override ;
        procedure OnArgument(S : String)     ; override ;
        procedure OnClose                    ; override ;
        procedure OnCancelRecord             ; override ;

        procedure OnClickMonnaieIn(Sender: TObject);
{$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          procedure OnClickChancellerie(Sender: TObject);
          procedure OnClickChancellerieOut(Sender: TObject);
        {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
        procedure OnChangeD_PARITEEURO;
    private
        OldDec,OldQuot : Integer;
        OldParite : double;
        Favertir,AvecMvt, bFromAss : Boolean;
{$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          MemoDev : String3;
        {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

        // lancé le fenetre de chancellerie !
        AsInsert : boolean;

        // Control du form
        D_MONNAIEIN,
        D_FONGIBLE,
        DEVPIVOT : TCheckBox;
        D_CPTLETTRDEBIT,
        D_CPTLETTRCREDIT,
        D_CPTPROVDEBIT,
        D_CPTPROVCREDIT,
        D_QUOTITE,
        D_PARITEEURO : THEdit; //: THDBEdit;
        TD_CPTLETTRDEBIT,
        TD_CPTLETTRCREDIT,
        TD_CPTPROVDEBIT,
        TD_CPTPROVCREDIT,
        TTD_CPTLETTRDEBIT,
        TTD_CPTLETTRCREDIT,
        TTD_CPTPROVDEBIT,
        TTD_CPTPROVCREDIT,
        DateDebutEuro : THLabel;
        D_DECIMALE : THSpinEdit; //: THDBSpinEdit;
        TSREGUL,
        TSEURO : TTabSheet;
        BChancel,
        BChancelOut,
        BValider : TSpeedButton;

        // Fonction
        function  GetCompteLibelle(cpt: string): CptEtLibelle;
        procedure PosDevPivot;
        function  VerifPariteEuro : boolean;
        Procedure EstMouvementer(St : String);
        Function  MouvementDepuisEuro(St : String) : boolean;
        procedure RechargeTablettes;
    end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
Function YYLanceFiche_DEVISE( vStRange, vStLequel, vStArgs : String ) : String ;
Procedure FicheDevise( vStCode : String ; vAction : TActionFiche ; vBoFromA : Boolean ) ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

//================================================================================
// Implementation
//================================================================================
implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
    paramsoc,
    wCommuns,
    {$IFDEF GCGC}
      EntGC,
    {$ENDIF GCGC}
    Hcompte,
{ GC_20080909_JTR_010;13258_Début }
    HTB97
{ GC_20080909_JTR_010;13258_Fin }
    ;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 18/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Function YYLanceFiche_DEVISE(vStRange,vStLequel,vStArgs : String ) : String;
begin
    Result := AGLLanceFiche('YY','DEVISE',vStRange,vStLequel,vStArgs);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 18/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Procedure FicheDevise(vStCode : String ; vAction : TActionFiche ; vBoFromA : Boolean);
var
    lStArgs : String;
begin
    case vAction of
        taConsult          : lStArgs := 'ACTION=CONSULTATION';
        taCreat,taCreatOne : lStArgs := 'ACTION=CREATION';
        else lStArgs := 'ACTION=MODIFICATION';
    end ;
    if (vBOFromA) then lStArgs := lStArgs + 'ASSISTANT=true';

    YYLanceFiche_DEVISE('',vStCode,lStArgs);
end ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

//================================================================================
// TOM_DEVISE
//================================================================================
//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnNewRecord;
begin
    inherited;
    SetField('D_DECIMALE',2) ;
    SetField('D_QUOTITE',1);
    SetField('D_FONGIBLE','-');
    SetField('D_MONNAIEIN','-');
    SetField('D_PARITEEURO',1);
    SetField('D_PARITEEUROFIXING', 1.00);
    SetField('D_ARRONDIPRIXACHAT', ArrondiPrecToNbDec(tAPCentieme));
    SetField('D_ARRONDIPRIXVENTE', ArrondiPrecToNbDec(tAPCentieme));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... : 01/10/2004
Description .. : JP 01/10/04 : gestion du Lasterror, car le delete était
Suite ........ : exécuté malgré le message d'interdiction de suppression
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnDeleteRecord;
begin
  if (VH^.TenueEuro and (GetField('D_DEVISE') = V_PGI.DeviseFongible)) then
  begin
    LastError := -1;
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer la devise pivot ou nationale.');
    Exit;
  end;

  try
      SourisSablier;
      EstMouvementer(GetField('D_DEVISE'));

      if (AvecMvt) then begin
        LastError := -1;
        LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer cette devise. Il existe des écritures, des pièces ou des tiers avec cette devise.');
        Exit;
      end;

      ExecuteSql('DELETE FROM CHANCELL WHERE H_DEVISE="'+GetField('D_DEVISE')+'"');

      inherited;

      RechargeTablettes; {JP 07/10/04}
      Favertir := true;
  finally
      SourisNormale;
  end;
  {$IFDEF GCGC}
    if assigned(VH_GC.TobDevise) then VH_GC.TobDevise.ClearDetail;
  {$ENDIF GCGC}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnUpdateRecord;
begin
  inherited;
  if (ds.State in [dsInsert]) then AsInsert := true;
  {$IFDEF COMPTA}
// ajout me pour la sychronisation
  MAJHistoparam ('DEVISE', GetField('D_DEVISE'));
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnAfterUpdateRecord;
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
var
    QLoc : TQuery;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
begin
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
    if (AsInsert = true) then
    begin
        if ((GetField('D_DEVISE') <> V_PGI.DevisePivot) and (not D_FONGIBLE.Checked)) then
        begin
            RechargeTablettes;{JP 07/10/04}
            if (not D_MONNAIEIN.Checked) then
            begin
                FicheChancel(GetField('D_DEVISE'),true,DebutDeMois(Date),taCreat,true);
//                AGLLanceFiche('CP','CPCHANCEL',GetField('D_DEVISE'),'','ACTION=CREATION;SUREURO');

                QLoc := OpenSql('SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE="'+GetField('D_DEVISE')+'"',true);
                if (QLoc.Eof) then HShowMessage(msg9, '', '')
                else if(QLoc.Fields[0].AsFloat=0) then HShowMessage(msg8, '', '');

                Ferme(QLoc) ;
            end;
        end;
        MemoDev := GetField('D_DEVISE');
        AsInsert := false;
    end;

    inherited;

    Favertir := true;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
  {$IFDEF GCGC}
    if assigned(VH_GC.TobDevise) then VH_GC.TobDevise.ClearDetail;
  {$ENDIF GCGC}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnLoadRecord;
begin
    inherited;

    OldDec := GetField('D_DECIMALE');
    OldQuot := GetField('D_QUOTITE');
    OldParite := GetField('D_PARITEEURO');

    // 10116
    SetControlEnabled('D_MAXDEBIT', True);
    SetControlEnabled('D_MAXCREDIT', True);
    if (GetField('D_DEVISE') = V_PGI.DevisePivot) then begin
      SetControlEnabled('D_MAXDEBIT', False);
      SetControlEnabled('D_MAXCREDIT', False);
    end;

    if ((D_MONNAIEIN.Checked) or (GetField('D_DEVISE') = V_PGI.DevisePivot)) then
    begin
        D_CPTLETTRDEBIT.Enabled := false;
        D_CPTLETTRCREDIT.Enabled := false;
        D_CPTPROVDEBIT.Enabled := false;
        D_CPTPROVCREDIT.Enabled := false;
        TD_CPTLETTRDEBIT.Enabled := false;
        TD_CPTLETTRCREDIT.Enabled := false;
        TD_CPTPROVDEBIT.Enabled := false;
        TD_CPTPROVCREDIT.Enabled := false;

        D_CPTLETTRDEBIT.Color := clBtnFace;
        D_CPTLETTRCREDIT.Color := clBtnFace;
        D_CPTPROVDEBIT.Color := clBtnFace;
        D_CPTPROVCREDIT.Color := clBtnFace;

        BChancel.Enabled := false;
    end
    else
    begin
        D_CPTLETTRDEBIT.Enabled := true;
        D_CPTLETTRCREDIT.Enabled := true;
        D_CPTPROVDEBIT.Enabled := true;
        D_CPTPROVCREDIT.Enabled := true;
        TD_CPTLETTRDEBIT.Enabled := true;
        TD_CPTLETTRCREDIT.Enabled := true;
        TD_CPTPROVDEBIT.Enabled := true;
        TD_CPTPROVCREDIT.Enabled := true;

        D_CPTLETTRDEBIT.Color := clWindow;
        D_CPTLETTRCREDIT.Color := clWindow;
        D_CPTPROVDEBIT.Color := clWindow;
        D_CPTPROVCREDIT.Color := clWindow;

        BChancel.Enabled := true;
    end;

    OnClickMonnaieIn(nil);

    if (D_FONGIBLE.Checked and VH^.TenueEuro) then
    begin
        D_PARITEEURO.Enabled := false;
        D_FONGIBLE.Enabled := false ;
    end;

    THLabel(GetControl('HL3',true)).Visible := D_PARITEEURO.Enabled;
    THLabel(GetControl('HL3',true)).Caption := TraduireMemoire(msg16) + ' ' + GetField('D_SYMBOLE') + ' ' + GetField('D_DEVISE') + ')' ;

    BChancelOut.Visible := (not D_MONNAIEIN.Checked);
    //SG6 13/01/05 FQ 15132 Suppression de visualisaton de la chancelerie en fr ou devise pivot pour les devise out
    BChancel.Visible := D_MONNAIEIN.Checked;
    BChancel.Enabled := (not (ds.State in [dsEdit,dsInsert]));
    BChancelOut.Enabled := (not (ds.State in [dsEdit,dsInsert]));

    if (GetField('D_FERME') = 'X') then
    begin
        BChancel.Enabled := false;
        BChancelOut.Enabled := false;
    end
    else
    begin
        BChancel.Enabled := true;
        BChancelOut.Enabled := true;
    end;

    if (GetField('D_DEVISE') = V_PGI.DevisePivot) then
    begin
        if (VH^.TenueEuro) then
        begin
            TSEURO.TabVisible := false;
            BChancelOut.Enabled := False;
        end
        // else TSEURO.TabVisible := true;
        { OL & CA le 06/07/2005 - si devise, onglet euro inutile }
        else
        begin
          TSEURO.TabVisible := False;
          BChancelOut.Enabled := False;
        end;
        BChancel.Enabled := False;
    end
//    else TSEURO.TabVisible := true;
    { OL & CA le 06/07/2005 - si devise, onglet euro inutile }
    else TSEURO.TabVisible := (GetField('D_MONNAIEIN')='X');

    PosDevPivot;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnChangeField(F: TField);
var
    cptLibelle: CptEtLibelle;
begin
    if (ds.state = dsBrowse) then exit;
    if (F.FieldName = 'D_FONGIBLE') then PosDevPivot;
    if (F.FieldName = 'D_PARITEEURO') then OnChangeD_PARITEEURO;

    BChancel.Enabled := not (ds.State in [dsEdit,dsInsert]);
    BChancelOut.Enabled := not (ds.State in [dsEdit,dsInsert]);

    if ((F.FieldName = 'D_CPTLETTRDEBIT') or (F.FieldName = 'D_CPTLETTRCREDIT') or (F.FieldName = 'D_CPTPROVDEBIT') or (F.FieldName = 'D_CPTPROVCREDIT')) then
    begin
        cptLibelle := GetCompteLibelle(GetField(F.FieldName));
        if (F.Asstring <> cptLibelle.Compte) then SetField(F.FieldName, cptLibelle.Compte);
        if (F.FieldName = 'D_CPTLETTRDEBIT') then TTD_CPTLETTRDEBIT.Caption := cptLibelle.libelle
        else if (F.FieldName = 'D_CPTLETTRCREDIT') then TTD_CPTLETTRCREDIT.Caption := cptLibelle.libelle
        else if (F.FieldName = 'D_CPTPROVDEBIT') then TTD_CPTPROVDEBIT.Caption := cptLibelle.libelle
        else TTD_CPTPROVCREDIT.Caption := cptLibelle.libelle;
    end;

    if ((F.FieldName = 'D_DECIMALE') and (ds.State in [dsEdit]) and (OldDec <> GetField('D_DECIMALE'))) then
    begin
        EstMouvementer(GetField('D_DEVISE'));
        if (AvecMvt) then
        begin
            HShowMessage(msg11, '', '');
            SetField('D_DECIMALE', OldDec);
            Abort;
            Exit;
        end;
    end;

    if ((F.FieldName = 'D_QUOTITE') and (ds.State in [dsEdit]) and (OldQuot <> GetField('D_QUOTITE'))) then
    begin
        EstMouvementer(GetField('D_DEVISE'));
        if (AvecMvt) then
        begin
            HShowMessage(msg12, '', '');
            SetField('D_QUOTITE', OldQuot);
            Abort;
            Exit;
        end;
    end;

    {JP 07/10/04 : not AfterInserting pour éviter d'avoir le message en insertion}
    if not AfterInserting and ((F.FieldName = 'D_PARITEEURO') and (not VerifPariteEuro)) then
    begin
        HShowMessage(msg14, '', '');
        Abort;
        Exit;
    end;

    if ((F.FieldName = 'D_PARITEEURO') and (ds.State in [dsEdit]) and (OldParite <> GetField('D_PARITEEURO')) and (not V_PGI.SAV)) then
    begin
        if (MouvementDepuisEuro(GetField('D_DEVISE'))) then
        begin
            HShowMessage(msg15, '', '');
            SetField('D_PARITEEURO', OldParite);
            Abort;
            Exit;
        end;
    end;
//  BBI fiche 13281
//  déplacé ici parce qu'en CWAS, le getfield('D_DEVISE') positionne le F.FieldName à 'D_DEVISE'

    if (GetField('D_DEVISE') = V_PGI.DevisePivot) then BChancel.Enabled := false;
    if (VH^.TenueEuro and (GetField('D_DEVISE') = V_PGI.DevisePivot)) then BChancelOut.Enabled := false;

//  BBI Fin fiche 13281
{$ifDEF SPEC302}
    if ((DEVPIVOT.Visible) and (DEVPIVOT.Checked) and (V_PGI.DevisePivot='') and bFromAss) then
    begin
        ExecuteSQL('UPDATE SOCIETE SET SO_DEVISEPRINC="' + GetField('D_DEVISE') + '"');
        {$IFNDEF EAGLSERVER} { => V_PGI.DevisePivot est en lecture seule en EAGLSERVER }
          V_PGI.DevisePivot := GetField('D_DEVISE');
        {$ENDIF !EAGLSERVIER}
    end;

    if (GetField('D_DEVISE') = V_PGI.DevisePivot) then
    begin
        V_PGI.OkDecV := GetField('D_DECIMALE');
        ExecuteSql('Update SOCIETE SET SO_DECVALEUR='+IntToStr(V_PGI.OkDecV)+', '+'SO_TAUXEURO='+StrfPoint(GetField('D_PARITEEURO'))+' '+'Where SO_SOCIETE="'+V_PGI.CodeSociete+'" AND SO_DEVISEPRINC="'+GetField('D_Devise')+'"');
    end;
{$else}
     if ((DEVPIVOT.Visible) and (DEVPIVOT.Checked) and (V_PGI.DevisePivot='') and bFromAss) then
     begin
        SetParamSoc('SO_DEVISEPRINC',GetField('D_DEVISE'));
        {$IFNDEF EAGLSERVER} { => V_PGI.DevisePivot est en lecture seule en EAGLSERVER }
          V_PGI.DevisePivot := GetField('D_DEVISE');
        {$ENDIF !EAGLSERVER}
     end;

     if (GetField('D_Devise') = V_PGI.DevisePivot) then
     begin
          {$IFNDEF EAGLSERVER} { => V_PGI.OkDecV est en lecture seule en EAGLSERVER  }
            V_PGI.OkDecV := GetField('D_DECIMALE');
          {$ENDIF !EAGLSERVER}
          SetParamSoc('SO_DECVALEUR', V_PGI.OkDecV);
          SetParamSoc('SO_TAUXEURO', GetField('D_PARITEEURO'));
     end;
{$endif}

    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnArgument(S : String);
var
    s1,s2 : string;
begin
    inherited;
    Ecran.HelpContext := 1150000;
    Favertir := false;
    bFromAss := false;
    AsInsert := false;

    s1 := uppercase(S);

    while (s1 <> '') do
    begin
        s2 := ReadTokenSt(s1);
        bFromAss := (s2 = 'ASSISTANT=true');
        if bFromAss then break;
    end;

    // recuperation des control
    DEVPIVOT := TCheckBox(GetControl('DEVPIVOT',true)); if (not assigned(DEVPIVOT)) then exit;
    D_MONNAIEIN := TCheckBox(GetControl('D_MONNAIEIN',true)); if (not assigned(D_MONNAIEIN)) then exit;
    D_FONGIBLE := TCheckBox(GetControl('D_FONGIBLE',true)); if (not assigned(D_FONGIBLE)) then exit;
    D_CPTLETTRDEBIT := THEdit(GetControl('D_CPTLETTRDEBIT',true)); if (not assigned(D_CPTLETTRDEBIT)) then exit;
    D_CPTLETTRCREDIT := THEdit(GetControl('D_CPTLETTRCREDIT',true)); if (not assigned(D_CPTLETTRCREDIT)) then exit;
    D_CPTPROVDEBIT := THEdit(GetControl('D_CPTPROVDEBIT',true)); if (not assigned(D_CPTPROVDEBIT)) then exit;
    D_CPTPROVCREDIT := THEdit(GetControl('D_CPTPROVCREDIT',true)); if (not assigned(D_CPTPROVCREDIT)) then exit;

    TD_CPTLETTRDEBIT := THLabel(GetControl('TD_CPTLETTRDEBIT',true)); if (not assigned(TD_CPTLETTRDEBIT)) then exit;
    TD_CPTLETTRCREDIT := THLabel(GetControl('TD_CPTLETTRCREDIT',true)); if (not assigned(TD_CPTLETTRCREDIT)) then exit;
    TD_CPTPROVDEBIT := THLabel(GetControl('TD_CPTPROVDEBIT',true)); if (not assigned(TD_CPTPROVDEBIT)) then exit;
    TD_CPTPROVCREDIT := THLabel(GetControl('TD_CPTPROVCREDIT',true)); if (not assigned(TD_CPTPROVCREDIT)) then exit;

    TTD_CPTLETTRDEBIT := THLabel(GetControl('TTD_CPTLETTRDEBIT',true)); if (not assigned(TTD_CPTLETTRDEBIT)) then exit;
    TTD_CPTLETTRCREDIT := THLabel(GetControl('TTD_CPTLETTRCREDIT',true)); if (not assigned(TTD_CPTLETTRCREDIT)) then exit;
    TTD_CPTPROVDEBIT := THLabel(GetControl('TTD_CPTPROVDEBIT',true)); if (not assigned(TTD_CPTPROVDEBIT)) then exit;
    TTD_CPTPROVCREDIT := THLabel(GetControl('TTD_CPTPROVCREDIT',true)); if (not assigned(TTD_CPTPROVCREDIT)) then exit;

    D_QUOTITE := THEdit(GetControl('D_QUOTITE',true)); if (not assigned(D_QUOTITE)) then exit;
    D_DECIMALE := THSpinEdit(GetControl('D_DECIMALE',true)); if (not assigned(D_DECIMALE)) then exit;
    D_PARITEEURO := THEdit(GetControl('D_PARITEEURO',true)); if (not assigned(D_PARITEEURO)) then exit;
    DateDebutEuro := THLabel(GetControl('DATEDEBUTEURO',true)); if (not assigned(DateDebutEuro)) then exit;

    TSREGUL := TTabSheet(GetControl('TSREGUL',true)); if (not assigned(TSREGUL)) then exit;
    TSEURO := TTabSheet(GetControl('TSEURO',true)); if (not assigned(TSEURO)) then exit;

    BChancel := TSpeedButton(GetControl('BCHANCEL',true)); if (not assigned(BChancel)) then exit;
    BChancelOut := TSpeedButton(GetControl('BCHANCELOUT',true)); if (not assigned(BChancelOut)) then exit;
    BValider := TSpeedButton(GetControl('BVALIDER',true)); if (not assigned(BValider)) then exit;

    {$IFDEF TRESO}
    {JP 20/08/03 : Dans la tréso, on ne s'occupe pas de la devise fongible}
    BChancel.Visible := False;
    BChancelOut.Left := BChancel.Left;
    {$ENDIF}

    // set des fonction
    D_MONNAIEIN.OnClick := OnClickMonnaieIn;
{$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      BChancel.OnClick := OnClickChancellerie;
      BChancelOut.OnClick := OnClickChancellerieOut;
    {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
//    BValider.OnClick := OnClickBValider;
    {JP 20/08/03 : Si l'on vient de créer une devise et que l'on va en Chancellerie}
{$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
	    TFFicheListe(ecran).AvertirTablettes := '';//'CPDEVISE;TTDEVISE;TTDEVISETOUTES;';
    {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

    // $$$ JP 16/05/06 - FQ BUREAU 11057
{$IFDEF BUREAU}
    tTabSheet(GetControl('TSFIXING', True)).TabVisible := FALSE;
{$ELSE}
    tTabSheet(GetControl('TSFIXING', True)).TabVisible := (GetParamSocSecur('SO_PREFSYSTTARIF',False) and  { FQ COMPTA 17068 }
                                                          (not (ctxCompta in V_PGI.PGIContexte)) and
                                                          (not (ctxTreso in V_PGI.PGIContexte))); {FQ TRESO 10211}
{$ENDIF}
{ GC_20080909_JTR_010;13258_Début }
    TToolbarButton97(GetControl('BTAUXNEG')).Visible := GetParamSocSecur('SO_GCTAUXNEGOCIE', false);
{ GC_20080909_JTR_010;13258_Fin }
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnClose;
begin
    if (FAvertir) then
      ChargeSocieteHalley;
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnCancelRecord;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnClickMonnaieIn(Sender: TObject);
var
    b: boolean;
begin
    b := ExisteSql('SELECT D_DEVISE FROM DEVISE WHERE D_FONGIBLE="X" AND D_DEVISE<>"'+ GetField('D_DEVISE') +'"');
    D_FONGIBLE.Enabled := D_MONNAIEIN.Checked and (not b);

    if (not (D_MONNAIEIN.Checked)) then
    begin
        if (D_FONGIBLE.Checked) then D_FONGIBLE.Checked := false;
        DateDebutEuro.Caption := '';
    end
    else
    begin
        if (not D_FONGIBLE.Enabled and D_FONGIBLE.Checked) then D_FONGIBLE.Checked := false;
        DateDebutEuro.Caption:=DateToStr(V_PGI.DateDebutEuro);
    end;

    D_PARITEEURO.Enabled := ((D_MONNAIEIN.Checked) and (V_PGI.DateDebutEuro >= EncodeDate(1999,01,01)) and (V_PGI.DateDebutEuro < EncodeDate(1999,12,31)));
    if (V_PGI.SAV) then D_PARITEEURO.Enabled := true;

    DateDebutEuro.Enabled := D_MONNAIEIN.Checked;
    THLabel(GetControl('TD_PARITEEURO',true)).Enabled := D_MONNAIEIN.Checked;
    THLabel(GetControl('HL4',true)).Enabled := D_MONNAIEIN.Checked;
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnClickChancellerie(Sender: TObject);
begin
  if (FAvertir) then AvertirTable('ttDevise');
  //    AGLLanceFiche('CP','CPCHANCELL','',GetField('D_DEVISE'),'ACTION=MODIFICATION;' + GetField('D_DEVISE') + ';PASEURO');
    FicheChancel(GetField('D_DEVISE'), True, DebutDeMois(Date), taModif, False);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnClickChancellerieOut(Sender: TObject);
begin
    if (FAvertir) then
    begin
        AvertirTable('ttDevise');
        AvertirTable('ttDeviseEtat');
    end;
//    AGLLanceFiche('CP','CPCHANCELL','',GetField('D_DEVISE'),'ACTION=MODIFICATION;' + GetField('D_DEVISE') + ';SUREURO');
  FicheChancel(GetField('D_DEVISE'), True, DebutDeMois(Date), taModif, True);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.OnChangeD_PARITEEURO;
var
    St : String ;
begin
    St := GetField('D_PARITEEURO');
    if (Valeur(St) > 999999.9) then
    begin
        HShowMessage(msg14, '', '');
        while (Valeur(St)>999999.9) do Delete(St,Length(St),1);
        SetField('D_PARITEEURO', St);
    end;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_DEVISE.GetCompteLibelle(cpt: string): CptEtLibelle;
var
    s,c,l : string;
    q : TQuery;
begin
    s := trim(cpt);

    if (s = '') then
    begin
        result.compte := '';
        result.libelle := '';
        exit;
    end;

    q := OpenSQL('SELECT G_GENERAL, G_LIBELLE From GENERAUX WHERE G_GENERAL Like "' + s +'%" ' + RecupWhere(tzGNonCollectif),true);

    if (not q.eof) then
    begin
        c := q.Fields[0].AsString;
        l := q.Fields[1].AsString;
        result.compte := c;
        result.libelle := l;
    end;

    Ferme(q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_DEVISE.VerifPariteEuro : boolean;
var
    StParite : String;
    Parite   : Integer;
begin
    Result := true;
    StParite := Trim(D_PARITEEURO.Text);

    if (Pos(V_PGI.SepDecimal,StParite) <> 0) then Delete(StParite,Pos(V_PGI.SepDecimal,StParite),1);

    Parite := StrToInt(StParite);

    if (Parite <= 0) then
    begin
        Result := false;
        Exit;
    end;

    StParite := IntToStr(Parite);

    if (Length(StParite) > 6) then Result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.PosDevPivot;
begin
    DEVPIVOT.Visible := ((D_FONGIBLE.Checked) and (V_PGI.DevisePivot='') and bFromAss);
    if (not DEVPIVOT.Visible) then DEVPIVOT.Checked := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DEVISE.EstMouvementer(St: String);
begin
    AvecMvt := ExisteSql('Select D_DEVISE from DEVISE Where D_DEVISE="'+St+'" '+'AND (Exists(Select E_DEVISE from ECRITURE Where E_DEVISE="'+St+'"))');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_DEVISE.MouvementDepuisEuro(St: String): boolean;
begin
  Result := ExisteSQL('SELECT E_DEVISE FROM ECRITURE WHERE E_DATECOMPTABLE>="'+UsDateTime(V_PGI.DateDebutEuro)+'" AND E_EXERCICE="'+QuelExoDT(V_PGI.DateDebutEuro)+'" AND E_DEVISE="'+St+'"');
end;

{JP 07/10/04 : Pour le rechargement des tablettes après une suppression ou une création
{---------------------------------------------------------------------------------------}
procedure TOM_DEVISE.RechargeTablettes;
{---------------------------------------------------------------------------------------}
begin
  AvertirTable('TTDEVISE');
  AvertirTable('CPDEVISE');
  AvertirTable('TTDEVISETOUTES');
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOM_DEVISE]);
end.
