{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/06/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPRELANCENIVEAU ()
Mots clefs ... : TOF;CPRELANCENIVEAU
*****************************************************************}
Unit CPRELANCENIVEAU_TOF ;

Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MainEAGL,
{$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    Fe_Main,
{$ENDIF}
    uTob,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    AGLInit,
    ULibEcriture, // WhereEcritureTob
    HRichOLE,     // THRichEditOLE, RichToString     
    Vierge,
    UTOF
    ;

function CPLanceFiche_RelanceNiveau ( vStRange, vStLequel, vStArgs : String ) : boolean;

Type
    TOF_CPRELANCENIVEAU = Class (TOF)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        nbEche : integer;
        { FQ 21116 BVE 17.07.07 }
        isOK   : boolean;
        { END FQ 21116 }
    end;

Implementation
uses HTB97;


function CPLanceFiche_RelanceNiveau(vStRange, vStLequel, vStArgs : String ) : boolean;
begin
    result := StrToBool(AGLLanceFiche('CP','CPRELANCENIVEAU',vStRange,vStLequel,vStArgs));
end ;

procedure TOF_CPRELANCENIVEAU.OnNew;
begin
  Inherited;
end;

procedure TOF_CPRELANCENIVEAU.OnDelete;
begin
  Inherited;
end;

procedure TOF_CPRELANCENIVEAU.OnUpdate;
begin
    Inherited;
    isOK := true;
    // MAJ des données de la fiche appelante
    TOB(TheData).PutValue('NIVEAURELANCE',StrToInt(GetControlText('NIVEAURELANCE')));
    TOB(TheData).PutValue('BLOCNOTE',  RichToString( THRichEditOLE(GetControl('BLOCNOTE')) ) ); // FQ 17060 SBO 05/12/2005
    // si on as modifier le bloc note et que l'on a plus d'une echeance alors !
    if (nbEche > 1) then
    begin
        if (HShowMessage('0;' + Ecran.Caption + ';Voulez-vous mettre à jour le bloc note de TOUTES les échéances selectionnées ?;Q;YN;N;N','','') = mrYes) then TOB(TheData).PutValue('MAJBLOCNOTE',true)
        else TOB(TheData).PutValue('MAJBLOCNOTE',false);
    end
    else TOB(TheData).PutValue('MAJBLOCNOTE',true);
end;

procedure TOF_CPRELANCENIVEAU.OnLoad;
begin
    Inherited;
end;

procedure TOF_CPRELANCENIVEAU.OnArgument(S : String);
begin
    Inherited;
    Ecran.HelpContext := 7568001 ;
    // recup des data
    SetControlText('NIVEAURELANCE',IntToStr(TOB(TheData).GetValue('NIVEAURELANCE')));
    SetControlText('BLOCNOTE',TOB(TheData).GetValue('BLOCNOTE'));
    // recup du nombre d'echeance a modifier
    nbEche := TOB(TheData).GetValue('NBECHE');
    isOK := false;
end;

procedure TOF_CPRELANCENIVEAU.OnClose;
begin
    Inherited;
    TFVierge(Ecran).Retour := BoolToStr(isOK);
end;

procedure TOF_CPRELANCENIVEAU.OnDisplay;
begin
    Inherited;
end;

procedure TOF_CPRELANCENIVEAU.OnCancel;
begin
    Inherited;
    isOk := false;
end;

Initialization
    registerclasses([TOF_CPRELANCENIVEAU]);
end.
