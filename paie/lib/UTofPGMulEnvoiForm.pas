{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGENVOIFORM_MUL ()
Mots clefs ... : TOF;PGENVOIFORM_MUL
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPGMulEnvoiForm ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,DBGrids,HDB,DBCtrls,Mul,
{$ELSE}
      mainEAGL,UtilEagl,EMul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HTB97,Hqry,ParamDat,UTob,PGOutilsFormation ;

Type
  TOF_PGENVOIFORM_MUL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    Q_Mul : THQuery;
    Arg : String;
    procedure AfficheEnvoi(Sender:TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure PurgeEnvoiForm(Sender : TObject);
  end ;

Implementation

procedure TOF_PGENVOIFORM_MUL.OnLoad ;
var Where : String;
begin
  Inherited ;
        Where := '';
        If StrtoInt(GetControlText('NUMENVOI'))>0 then Where := 'PVF_NUMENVOI='+GetControlText('NUMENVOI')+'';
        SetControlText('XX_WHERE',Where);
end ;

procedure TOF_PGENVOIFORM_MUL.OnArgument (S : String ) ;
var BDelete : TToolBarButton97;
    Date1,Date2 : THEdit;
    DateDebut,DateFin : TDateTime;
    Q : TQuery;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        {$ENDIF}
        if Grille <> NIL then Grille.OnDblClick := AfficheEnvoi;
        Date1 := THEdit(GetControl('PVF_PREPARELE'));
        If Date1 <> Nil then Date1.OnElipsisClick := DateElipsisClick;
        Date2 := THEdit(GetControl('PVF_PREPARELE_'));
        If Date2 <> Nil then Date2.OnElipsisClick := DateElipsisClick;
        RendMillesimeRealise(DateDebut,DateFin);
        SetcontrolText('PVF_PREPARELE',DateToStr(dateDebut));
        SetcontrolText('PVF_PREPARELE_',DateToStr(datefin));
        If Arg='ENVOI' then SetControlText('PVF_STATUTENVOI','005');
        If Arg='FACTURE' then SetControlText('PVF_STATUTENVOI','010');
        If Arg='PURGE' then
                begin
                SetControlVisible('BSELECTALL',True);
                SetControlProperty ('FListe','Multiselection', TRUE);
                SetControlVisible('BDELETE',True);
                BDelete := TToolBarButton97(GetControl('BDELETE'));
                If BDelete <> Nil then BDelete.OnClick := PurgeEnvoiForm;
                Ecran.Caption := 'Purge des envois de formation';
                UpdateCaption(Ecran);
                end;
end ;

procedure TOF_PGENVOIFORM_MUL.AfficheEnvoi(Sender : TObject);
var St : String;
    OPCA,Support,Reel,MessError,EmisPar : String;
begin
        st  := IntToStr(Q_MUL.FindField('PVF_NUMENVOI').AsInteger);
        If arg='MODIF' then AGLLanceFiche('PAY','ENVOIFORMATION','',St,Arg+';ACTION=MODIFICATION')
        Else
        begin
                If arg='PURGE' then AGLLanceFiche('PAY','ENVOIFORMATION','',St,Arg+';ACTION=CONSULTATION')
                Else
                begin
                        OPCA := GetControltext('PVF_ORGCOLLECTGU');
                        Support := GetControltext('SUPPORTEMIS');
                        Reel := GetControltext('ENVOIREEL');
                        EmisPar := GetControltext('EMETSOC');
                        MessError := '';
                        If OPCA='' then MessError := '#13#10- L''organisme collecteur';
                        If Support='' then MessError := MessError+'#13#10- Le support';
                        If EmisPar='' then MessError := MessError+'#13#10- L''émetteur';
                        If MessError<>'' then
                        begin
                                PGIBox('Vous devez renseigner '+MessError,Ecran.Caption);
                                Exit;
                        end;
                        AGLLanceFiche('PAY','ENVOIFORM','','',St+';'+OPCA+';'+Support+';'+Reel+';'+EmisPar);
                end;
        end;
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGENVOIFORM_MUL.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGENVOIFORM_MUL.PurgeEnvoiForm(Sender : TObject);
var     TobEnvoi,T : Tob;
        Q : TQuery;
        i,NumEnvoiSupp : Integer;
        DateDeb,DateFin : TDateTime;
begin
        DateDeb := StrToDate(GetControlText('PVF_PREPARELE'));
        DateFin := StrToDate(GetControlText('PVF_PREPARELE_'));
        Q := OpenSQL('SELECT * FROM ENVOIFORMATION WHERE PVF_PREPARELE>="'+UsDateTime(DateDeb)+'"'+
        ' AND PVF_PREPARELE<="'+UsDateTime(DateFin)+'"',True);
        TobEnvoi := Tob.Create('ENVOIFORMATION',Nil,-1);
        TobEnvoi.LoadDetailDB('ENVOIFORMATION','','',Q,False);
        Ferme(Q);
        If (Grille = nil) then Exit;
        if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;
        if ((Grille.nbSelected) > 0) AND (not Grille.AllSelected ) then
        begin
                for i := 0 to Grille.NbSelected-1 do
                begin
                        Grille.GotoLeBookmark(i);
                        {$IFDEF EAGLCLIENT}
                        TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
                        {$ENDIF}
                        NumEnvoiSupp := Q_mul.FindField('PVF_NUMENVOI').AsInteger;
                        T  :=  TobEnvoi.FindFirst(['PVF_NUMENVOI'],[NumEnvoiSupp],False);
                        If T<>Nil then
                        begin
                                T.DeleteDB(False);
                        end;
                end;
        Grille.ClearSelected;
        end;
         {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
         {$ENDIF}
        If (Grille.AllSelected=TRUE) then
        begin
                Q_mul.First;
                while Not Q_mul.EOF do
                begin
                        NumEnvoiSupp := Q_mul.FindField('PVF_NUMENVOI').AsInteger;
                        T := TobEnvoi.FindFirst(['PVF_NUMENVOI'],[NumEnvoiSupp],False);
                        If T<>Nil then
                        begin
                                T.DeleteDB(False);
                        end;
                        Q_mul.Next;
                end;
                Grille.AllSelected:=False;
        end;
        TobEnvoi.Free;
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGENVOIFORM_MUL ] ) ;
end.


