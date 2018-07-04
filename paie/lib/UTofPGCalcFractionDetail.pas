{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCALCFRACDETAIL ()
Mots clefs ... : TOF;PGCALCFRACDETAIL
*****************************************************************
PT1 11/08/2004 V_50 JL FQ 11484
}
Unit UTofPGCalcFractionDetail ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,HTB97,UTob,HSysMenu,Dialogs ;

Type
  TOF_PGCALCFRACDETAIL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    LaDate,LaFraction,TypeSaisie,Predefini,NumDos : String;
    Grille : THGrid;
    procedure DateElipsisclick(Sender : TObject);
    procedure UpdateFraction (Sender : TObject);
    procedure EffacerGrille (Sender : TObject);
    procedure RecupereDetail ;
    procedure Suppdetail (Sender : TObject);
  end ;

Implementation

procedure TOF_PGCALCFRACDETAIL.OnArgument (S : String ) ;
Var Edit : THEdit;
    BValid,BDelete,BEfface : TToolBarButton97;
    HMTrad: THSystemMenu;
    Q : TQuery;
begin
  Inherited ;
        TypeSaisie := ReadTokenPipe(S,';');
        LaFraction := ReadTokenPipe(S,';');
        LaDate := ReadTokenPipe(S,';');
        SetControlText('DATEVALID',DateToStr(Date));
        SetControlText('AUGACHARGE','0');
        Edit := THEdit(GetControl('DATEVALID'));
        If Edit <> NIL Then Edit.OnElipsisClick := DateElipsisClick;
        Ecran.Caption := 'Saisie calcul retenue sur salaire : '+RechDom('PGCALCFRACTION',LaFraction,False);
        Grille := THGrid(GetControl('GRETENUE'));
        Grille.ColAligns[0]  := taRightJustify;
        Grille.ColAligns[1]  := taRightJustify;
        Grille.ColAligns[2]  := TaCenter;
        Grille.ColAligns[3]  := TaCenter;
        Grille.ColFormats[0] := '# ##0.00';
        Grille.ColFormats[1] := '# ##0.00';
        Grille.ColFormats[2] := '##0';
        Grille.ColFormats[3] := '##0';
        BValid := TToolBarButton97(Getcontrol('BVALID'));
        if BValid <> Nil then BValid.OnClick := UpdateFraction;
        If TypeSaisie = 'MODIF' then
        begin
                If StrToDate(LaDate) > IDate1900 then SetControlText('DATEVALID',LaDate);
                RecupereDetail;
        end
        else HMTrad.ResizeGridColumns(Grille) ;
                BEfface := TToolBarButton97(GetControl('BEFFACE'));
        If BEfface <> Nil then BEfface.OnClick := EffacerGrille;
        BDelete := TToolBarButton97(GetControl('BDELETE'));
        If BDelete <> Nil then BDelete.OnClick := Suppdetail;
        Predefini := '';
        NumDos := '';
        Q := OpenSQL('SELECT PCF_NODOSSIER,PCF_PREDEFINI FROM CALCFRACTION WHERE PCF_CALCFRACTION="'+LaFraction+'"',True);//PT1
        If not Q.eof then
                begin
                        Predefini := Q.FindField('PCF_NODOSSIER').AsString;
                        NumDos := Q.FindField('PCF_NODOSSIER').AsString;
                end;
        Ferme(Q);
end ;

procedure TOF_PGCALCFRACDETAIL.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGCALCFRACDETAIL.UpdateFraction (Sender : TObject);
var Numerateur,Denominateur,i,MaxOrdre : Integer;
    TobFraction,TobMaj : Tob;
    Q : Tquery;
    MontantInf,MontantSup,AugACharge : Double;
    DateValid : TDateTime;
begin
        MaxOrdre := 1;
        If Grille = Nil then Exit;
        DateValid := StrToDate(GetControlText('DATEVALID'));
        If GetControlText('AUGACHARGE') = '' then
        begin
                PGIBox('Vous devez renseigner l''augmentation par personne à charge',ecran.Caption);
                Exit;
        end;
        AugACharge := StrToFloat(GetControlText('AUGACHARGE'));
        If TypeSaisie='MODIF' then ExecuteSQL('DELETE FROM CALCFRACTIONDETAIL WHERE '+
        'PCD_CALCFRACTION="'+LaFraction+'" AND PCD_DATEDEBUT="'+UsDateTime(StrToDate(LaDate))+'"');
        For i := 1 to Grille.RowCount - 1 do
        begin
                If Grille.CellValues[0,i] <> '' then MontantInf := STrToFloat(Grille.CellValues[0,i])
                else MontantInf := 0;
                If Grille.CellValues[0,i] <> '' then MontantSup := STrToFloat(Grille.CellValues[1,i])
                else MontantSup := 0;
                If Grille.CellValues[0,i] <> '' then Numerateur := STrToInt(Grille.CellValues[2,i])
                else Numerateur := 0;
                If Grille.CellValues[0,i] <> '' then Denominateur := STrToInt(Grille.CellValues[3,i])
                else Denominateur := 1;
                If (MontantInf = 0) and (MontantSup = 0) then Exit;
                If Numerateur = 0 then exit;
                TobFraction := Tob.Create('CALCFRACTIONDETAIL',Nil,-1);
                TobFraction.PutValue('PCD_CALCFRACTION',LaFraction);
                TobFraction.PutValue('PCD_ORDRE',MaxOrdre);
                TobFraction.PutValue('PCD_DATEDEBUT',DateValid);
                TobFraction.PutValue('PCD_DATEFIN',IDate1900);
                TobFraction.PutValue('PCD_MONTANTINF',MontantInf);
                TobFraction.PutValue('PCD_MONTANTSUP',MontantSup);
                TobFraction.PutValue('PCD_NUMERATEUR',Numerateur);
                TobFraction.PutValue('PCD_DENOMINATEUR',Denominateur);
                TobFraction.PutValue('PCD_NBJOURS',0);
                TobFraction.PutValue('PCD_AUGACHARGE',AugACharge);
                TobFraction.PutValue('PCD_NODOSSIER',NumDos);
                TobFraction.PutValue('PCD_PREDEFINI',Predefini);
                TobFraction.InsertOrUpdateDB(False);
                MaxOrdre := MaxOrdre + 1;
                TobFraction.Free;
        end;
end;

procedure TOF_PGCALCFRACDETAIL.EffacerGrille (Sender : TObject);
var i : Integer;
    GrilleRem : THGrid;
begin
        GrilleRem := THGrid(GetControl('GRETENUE'));
        For i := 1 to GrilleRem.RowCount -1 do
        begin
                GrilleRem.Rows[i].Clear;
        end;
        GrilleRem.RowCount := 10;
end;

procedure TOF_PGCALCFRACDETAIL.RecupereDetail ;
var Q : TQuery;
    TobRem,TobGrille,T : Tob;
    StInf,StSup : String;
    GrilleRem : THGrid;
    i : Integer;
    MontantInf,MontantSup : Double;
    Numerateur,Denominateur : Integer;
    HMTrad: THSystemMenu;
begin
        GrilleRem := THGrid(GetControl('GRETENUE'));
        For i := 1 to GrilleRem.RowCount -1 do
        begin
                GrilleRem.Rows[i].Clear;
        end;
        GrilleRem.RowCount := 10;
        GrilleRem.ColAligns[0]:=taRightJustify;
        GrilleRem.ColAligns[1]:=taRightJustify;
        GrilleRem.ColAligns[2]:=taCenter;
        Q := OpenSQL('SELECT PCD_MONTANTINF,PCD_MONTANTSUP,PCD_NUMERATEUR,PCD_DENOMINATEUR FROM CALCFRACTIONDETAIL'+
        ' WHERE PCD_CALCFRACTION="'+LaFraction+'" AND PCD_DATEDEBUT="'+UsDateTime(StrToDate(LaDate))+'" AND PCD_ORDRE>0 ORDER BY PCD_ORDRE',True);
        TobRem := Tob.Create('Les rémunérations',Nil,-1);
        TobRem.LoadDetailDB('Les rémunérations','','',Q,False);
        Ferme(Q);
        TobGrille := Tob.Create('Remplir Grille',Nil,-1);
        For i:= 0 to TobRem.detail.Count - 1 do
        begin
                MontantInf := TobRem.Detail[i].GetValue('PCD_MONTANTINF');
                StInf := FloatToStr(MontantInf);
                MontantSup := TobRem.Detail[i].Getvalue('PCD_MONTANTSUP');
                If MontantSup > 0 then StSup := FloatToStr(MontantSup)
                Else StSup := '/';
                Numerateur := TobRem.Detail[i].GetValue('PCD_NUMERATEUR');
                Denominateur := TobRem.Detail[i].GetValue('PCD_DENOMINATEUR');
                T := Tob.Create('Une fille',TobGrille,-1);
                T.AddChampSupValeur('INFERIEUR',StInf,False);
                T.AddChampSupValeur('SUPERIEUR',StSup,False);
                T.AddChampSupValeur('NUMERATEUR',IntToStr(Numerateur),False);
                T.AddChampSupValeur('DENOMINATEUR',IntToStr(Denominateur),False);
        end;
        TobRem.Free;
        TobGrille.PutGridDetail(GrilleRem,False,True,'',False);
        TobGrille.Free;
        HMTrad.ResizeGridColumns(GrilleRem) ;
        Q := OpenSQL('SELECT PCD_AUGACHARGE FROM CALCFRACTIONDETAIL'+
        ' WHERE PCD_CALCFRACTION="'+LaFraction+'" AND PCD_DATEDEBUT="'+UsDateTime(StrToDate(LaDate))+'" AND PCD_ORDRE=1 ORDER BY PCD_ORDRE',True);
        If Not Q.Eof then SetControlText('AUGACHARGE',FloatToStr(Q.FindField('PCD_AUGACHARGE').AsFloat));
        Ferme(Q);
end;

procedure TOF_PGCALCFRACDETAIL.Suppdetail (Sender : TObject);
begin
        If PGIAskCancel('Confimez-vous la suppression de l''enregistrement ?',Ecran.Caption) = MrYes then
        begin
                ExecuteSQL('DELETE FROM CALCFRACTIONDETAIL WHERE '+
                'PCD_CALCFRACTION="'+LaFraction+'" AND PCD_DATEDEBUT="'+UsDateTime(StrToDate(LaDate))+'"');
                Ecran.Close;
        end;
end;


Initialization
  registerclasses ( [ TOF_PGCALCFRACDETAIL ] ) ;
end.

