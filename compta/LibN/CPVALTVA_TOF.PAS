{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPVALTVA ()
Mots clefs ... : TOF;CPVALTVA
*****************************************************************}

Unit CPVALTVA_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGL,
    eMul,
{$ELSE}
    FE_Main,
    db,
    dbtables,
    Mul,
{$ENDIF}
    Saisie,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    Ent1,
    HMsgBox,
    UTOF,
    HTB97,
    HStatus,
    ParamSoc,		// GetParamSocSecur YMO
    SAISUTIL
    ;

//==================================================
// Externe
//==================================================
procedure TvaValidEnc;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPVALTVA = Class(TOF)
    public
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        procedure OnChangeE_EXERCICE(Sender : TObject);
        procedure OnClickBOuvrir(Sender: TObject);
        procedure OnClickBVALIDER(Sender: TObject);
        procedure OnClickBZOOMECRITURE(Sender: TObject);
        procedure OnDblClickFliste(Sender: TObject);
        procedure AuxiElipsisClick(Sender : TObject);

        procedure ShowEcriture;
        procedure UpdateValidEnc;
        procedure ValEditEnc;
    public
    private
        Fliste : THGrid;
        NowFutur : TDateTime;

    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

//==================================================
// Definition de variables
//==================================================
var
    Mess : array [0..2] of string = (
    {00}    'ATTENTION. Modification non enregistrée !',
    {01}    '1;?caption?;Confirmez-vous la validation des écritures éditées ?;Q;YN;Y;Y;',
    {02}    '2;?caption?;Vous n''avez sélectionné aucune écriture;W;O;O;O;'
    );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TvaValidEnc;
begin
    AGLLanceFiche('CP','CPVALTVA','','','');
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnNew;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnUpdate;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnLoad;
begin
    Inherited;

    SetControlText('XX_WHERE','(E_NUMECHE=1 AND E_QUALIFPIECE="N" AND E_TRESOLETTRE<>"X" AND E_NATUREPIECE<>"ECC" AND E_ECHE="X" AND E_EDITEETATTVA="X" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H"))');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnArgument(S : String);
var
    E_EXERCICE,E_JOURNAL : THValComboBox;
    BOuvrir,BVALIDER,BZOOMECRITURE : TToolBarButton97;
begin
    if (not VH^.OuiTvaEnc) then
    begin
        HShowMessage('10;Tva sur encaissements;Le module n''est pas installé !;E;O;O;O;','','');
        exit;
    end;

    Inherited;

    SetControlText('E_DEVISE',V_PGI.DevisePivot);

    E_EXERCICE := THValComboBox(GetControl('E_EXERCICE',true)); if (not assigned(E_EXERCICE)) then exit;
    E_EXERCICE.OnChange := OnChangeE_EXERCICE;
    E_JOURNAL := THValComboBox(GetControl('E_JOURNAL',true)); if (not assigned(E_JOURNAL)) then exit;

    if (VH^.CPExoRef.Code <> '') then
    begin
        E_EXERCICE.Value := VH^.CPExoRef.Code;
        SetControlText('E_DATECOMPTABLE',DateToStr(VH^.CPExoRef.Deb));
        SetControlText('E_DATECOMPTABLE_',DateToStr(VH^.CPExoRef.Fin));
    end
    else
    begin
        E_EXERCICE.Value := VH^.Entree.Code;
        SetControlText('E_DATECOMPTABLE',DateToStr(V_PGI.DateEntree));
        SetControlText('E_DATECOMPTABLE_',DateToStr(V_PGI.DateEntree));
    end;

    SetControlText('E_DATEECHEANCE',StDate1900);
    SetControlText('E_DATEECHEANCE_',StDate2099);
    SetControlText('E_DATECREATION',StDate1900);
    SetControlText('E_DATECREATION_',StDate2099);

    PositionneEtabUser(THValComboBox(GetControl('E_ETABLISSEMENT',true)));

    if ((E_JOURNAL.Value = '') and (E_JOURNAL.Values.Count > 0)) then
    begin
        if (not E_JOURNAL.Vide) then E_JOURNAL.ItemIndex := 0
        else if (E_JOURNAL.Values.Count > 1) then E_JOURNAL.ItemIndex := 1;
    end;

    // recup des btns pour ajouté les fcts
    BOuvrir := TToolBarButton97(GetControl('BOuvrir',true));  if (not assigned(BOuvrir)) then exit;
    BOuvrir.OnClick := OnClickBOuvrir;
    BVALIDER := TToolBarButton97(GetControl('BVALIDER',true));  if (not assigned(BVALIDER)) then exit;
    BVALIDER.OnClick := OnClickBVALIDER;
    BZOOMECRITURE := TToolBarButton97(GetControl('BZOOMECRITURE',true));  if (not assigned(BZOOMECRITURE)) then exit;
    BZOOMECRITURE.OnClick := OnClickBZOOMECRITURE;
    // idem grille
    Fliste := THGrid(GetControl('Fliste',true));  if (not assigned(Fliste)) then exit;
    Fliste.OnDblClick := OnDblClickFliste;

    if GetParamSocSecur('SO_CPMULTIERS', false) then
        THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnClose;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnCancel;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnChangeE_EXERCICE(Sender : TObject);
begin
    ExoToDates(THValComboBox(Sender).Value,GetControl('E_DATECOMPTABLE',true),GetControl('E_DATECOMPTABLE_',true)) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnClickBVALIDER(Sender : TObject);
begin
    if ((not FListe.AllSelected) and (FListe.NbSelected <= 0)) then
    begin
        HShowMessage(Mess[2],ecran.caption,'');
        exit;
    end;

    if (HShowMessage(Mess[1],ecran.caption,'') <> mrYes) then exit;

    NowFutur := NowH;

    if (Transactions(ValEditEnc,2) <> oeOk) then MessageAlerte(Mess[0])
    else
    begin
        if (not FListe.AllSelected) then Fliste.ClearSelected
        else FListe.AllSelected := false;

        TFMul(Ecran).BChercheClick(Nil);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnClickBZOOMECRITURE(Sender : TObject);
begin
    ShowEcriture;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnClickBOuvrir(Sender : TObject);
begin
    ShowEcriture;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.OnDblClickFliste(Sender : TObject);
begin
    ShowEcriture;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.ValEditEnc;
var
    i : integer;
begin
    if (Fliste.AllSelected) then
    begin
        InitMove(100,'');

        TFMul(Ecran).Q.First;

        while (not TFMul(Ecran).Q.EOF) do
        begin
            MoveCur(False);
            UpdateValidEnc;
            if (V_PGI.IoError <> oeOk) then break;
            TFMul(Ecran).Q.Next;
        end;
    end
    else
    begin
        InitMove(FListe.NbSelected,'');

        for i := 0 to FListe.NbSelected-1 do
        begin
            MoveCur(False);
            FListe.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
            UpdateValidEnc;
            if (V_PGI.IoError <> oeOk) then break;
        end;
    end;

    FiniMove;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.UpdateValidEnc;
var
    Query : String;
    Nb : Integer;
begin
    Query := 'UPDATE ECRITURE SET E_EDITEETATTVA="#", E_DATEMODIF="' + USTime(NowFutur) + '" WHERE E_JOURNAL="' + TFMul(Ecran).Q.FindField('E_JOURNAL').AsString + '" AND E_EXERCICE="' + QuelExo(DateToStr(TFMul(Ecran).Q.FindField('E_DATECOMPTABLE').AsDateTime)) + '" AND E_DATECOMPTABLE="' + USDATETIME(TFMul(Ecran).Q.FindField('E_DATECOMPTABLE').AsDateTime) + '" AND E_NUMEROPIECE=' + TFMul(Ecran).Q.FindField('E_NUMEROPIECE').AsString + ' AND E_NUMLIGNE=' + TFMul(Ecran).Q.FindField('E_NUMLIGNE').AsString + ' AND E_NUMECHE>=1 AND E_QUALIFPIECE="N" AND E_EDITEETATTVA="X"';
    Nb := ExecuteSQL(Query);

    if (Nb <= 0) then V_PGI.IoError := oeUnknown;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALTVA.ShowEcriture;
begin
{$IFDEF EAGLCLIENT}
    if ((TFMul(Ecran).Q.TQ.EOF) And (TFMul(Ecran).Q.TQ.BOF)) then exit;
    TrouveEtLanceSaisie(TFMul(Ecran).Q.TQ,taConsult,GetControlText('E_QUALIFPIECE'));
{$ELSE}
    if ((TFMul(Ecran).Q.EOF) And (TFMul(Ecran).Q.BOF)) then exit;
    TrouveEtLanceSaisie(TFMul(Ecran).Q,taConsult,GetControlText('E_QUALIFPIECE'));
{$ENDIF}
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPVALTVA]);
end.
