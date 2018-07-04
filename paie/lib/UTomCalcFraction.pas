{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CALCFRACTION (CALCFRACTION)
Mots clefs ... : TOM;CALCFRACTION
*****************************************************************
PT1 03/01/2005 JL V_60 Accès bouton création périodes
}
Unit UTomCalcFraction ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,Fiche,
{$ELSE}
     MainEAGL,eFiche,eFichList,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOM,UTob,HTB97,HSysMenu,PGOutils ;

Type
  TOM_CALCFRACTION = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
     Arg : String;
     CEG,LectureSeule,STD,DOS : Boolean;
    procedure CreerDetail (Sender : TObject);
    procedure AfficherGrilleRetenue ;
    procedure ChangePeriodeRetenue (Sender : TObject);
    procedure GrilleRetenueDblClick (Sender : TObject);
    end ;

Implementation

procedure TOM_CALCFRACTION.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_CALCFRACTION.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CALCFRACTION.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CALCFRACTION.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CALCFRACTION.OnLoadRecord ;
var Q : TQuery;
    TobDetail : Tob;
    combo : THValComboBox;
    i : Integer;
    ValeurCombo : String;
begin
  Inherited ;
        AccesPredefini('TOUS',CEG,STD,DOS);
        LectureSeule := False;
        if (Getfield('PCF_PREDEFINI') = 'CEG')  then
        begin
                LectureSeule := (CEG=False);
                PaieLectureSeule(TFFiche(Ecran),(CEG=False));
                SetControlEnabled('BDETAIL',CEG);  // PT1
                SetControlVisible('BDETAIL',CEG);  // PT1
                SetControlEnabled('BDelete',CEG);
        end;
        Combo := THValComboBox(GeTcontrol('VALIDITE'));
        Q := OpenSQL('SELECT PCD_DATEDEBUT FROM CALCFRACTIONDETAIL '+
        'WHERE PCD_ORDRE=1 AND PCD_CALCFRACTION="'+GetField('PCF_CALCFRACTION')+'" ORDER BY PCD_DATEDEBUT DESC',True);
        TobDetail := Tob.Create('Les fractions',Nil,-1);
        TobDetail.LoadDetailDB('CALCFRACTIONDETAIL','','',Q,False);
        Ferme(Q);
        Combo.Items.Clear;
        Combo.Values.Clear;
        ValeurCombo := '';
        For i := 0 to TobDetail.Detail.Count -1 do
        begin
                if i = 0 then ValeurCombo := DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT'));
                Combo.Items.Add('à partir du '+DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT')));
                Combo.Values.Add(DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT')));
        end;
        Combo.Value := ValeurCombo;
        AfficherGrilleRetenue;
        if Arg = 'RETENUESAL' then
        begin
                SetField('PCF_TYPECALCFRAC','001');
                SetControlEnabled('PCF_TYPECALCFRAC',False);
        end;
end ;

procedure TOM_CALCFRACTION.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_CALCFRACTION.OnArgument ( S: String ) ;
var Combo : THValComboBox;
    BDetail : TToolBarButton97;
    Grille : THGrid;

begin
  Inherited ;
        Arg := ReadTokenPipe (S,';');
        if Arg = 'RETENUESAL' then
        begin
                SetControlProperty('PMAINTIEN','TabVisible',False);
                SetControlVisible('PMAINTIEN',False);
        end;
        Combo := THValComboBox(GeTcontrol('VALIDITE'));
        If Combo <> Nil then Combo.OnClick := ChangePeriodeRetenue;
        BDetail := TToolBarButton97(GetControl('BDETAIL'));
        If BDetail <> Nil then BDetail.OnClick := CreerDetail;
        Grille := THGrid(GetControl('GRETENUE'));
        If Grille <> Nil then Grille.OnDblClick := GrilleRetenueDblClick;
end ;

procedure TOM_CALCFRACTION.CreerDetail (Sender : TObject);
var St : String;
    Q : TQuery;
    TobDetail : Tob;
    combo : THValComboBox;
    i : Integer;
    ValeurCombo : String;
begin
        St := GetField('PCF_CALCFRACTION')+';01/01/1900';
        AGLLanceFiche('PAY','CALCFRACDETAIL','','','CREATION;'+St);
        Combo := THValComboBox(GeTcontrol('VALIDITE'));
        Q := OpenSQL('SELECT PCD_DATEDEBUT FROM CALCFRACTIONDETAIL '+
        'WHERE PCD_ORDRE=1 AND PCD_CALCFRACTION="'+GetField('PCF_CALCFRACTION')+'" ORDER BY PCD_DATEDEBUT DESC',True);
        TobDetail := Tob.Create('Les fractions',Nil,-1);
        TobDetail.LoadDetailDB('CALCFRACTIONDETAIL','','',Q,False);
        Ferme(Q);
        Combo.Items.Clear;
        Combo.Values.Clear;
        ValeurCombo := '';
        For i := 0 to TobDetail.Detail.Count -1 do
        begin
                if i = 0 then ValeurCombo := DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT'));
                Combo.Items.Add('à partir du '+DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT')));
                Combo.Values.Add(DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT')));
        end;
        Combo.Value := ValeurCombo;
        AfficherGrilleRetenue;
end;

procedure TOM_CALCFRACTION.AfficherGrilleRetenue ;
var Q : TQuery;
    TobRem,TobGrille,T : Tob;
    StDate,StInf,StSup : String;
    GrilleRem : THGrid;
    i : Integer;
    MontantInf,MontantSup : Double;
    Numerateur,Denominateur : Integer;
    HMTrad: THSystemMenu;
    Combo : THValComboBox;
begin
        GrilleRem := THGrid(GetControl('GRETENUE'));
        For i := 1 to GrilleRem.RowCount -1 do
        begin
                GrilleRem.Rows[i].Clear;
        end;
        Combo := THValComboBox(GetControl('VALIDITE'));
        GrilleRem.RowCount := 2;
        GrilleRem.ColAligns[0]:=taRightJustify;
        GrilleRem.ColAligns[1]:=taRightJustify;
        GrilleRem.ColAligns[2]:=taCenter;
        StDate := GetControlText('VALIDITE');
        If not IsValidDate(StDate) then Exit;
        Q := OpenSQL('SELECT PCD_MONTANTINF,PCD_MONTANTSUP,PCD_NUMERATEUR,PCD_DENOMINATEUR FROM CALCFRACTIONDETAIL'+
        ' WHERE PCD_CALCFRACTION="'+GetField('PCF_CALCFRACTION')+'" AND PCD_DATEDEBUT="'+UsDateTime(StrToDate(Combo.Value))+'" AND PCD_ORDRE>0 ORDER BY PCD_ORDRE',True);
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
                T.AddChampSupValeur('FRACTION',IntToStr(Numerateur)+'/'+IntToStr(Denominateur),False);
        end;
        TobRem.Free;
        TobGrille.PutGridDetail(GrilleRem,False,True,'',False);
        TobGrille.Free;
        Q := OpenSQL('SELECT PCD_AUGACHARGE FROM CALCFRACTIONDETAIL'+
        ' WHERE PCD_CALCFRACTION="'+GetField('PCF_CALCFRACTION')+'" AND PCD_DATEDEBUT="'+UsDateTime(StrToDate(Combo.Value))+'" AND PCD_ORDRE=1 ORDER BY PCD_ORDRE',True);
        If Not Q.Eof then SetControlText('AUGACHARGE',FloatToStr(Q.FindField('PCD_AUGACHARGE').AsFloat));
        Ferme(Q);
        HMTrad.ResizeGridColumns(GrilleRem) ;
end;

procedure TOM_CALCFRACTION.ChangePeriodeRetenue (Sender : TObject);
begin
        AfficherGrilleRetenue;
end;

procedure TOM_CALCFRACTION.GrilleRetenueDblClick (Sender : TObject);
var St : String;
    Q : TQuery;
    TobDetail : Tob;
    combo : THValComboBox;
    i : Integer;
    ValeurCombo : String;
begin
        St := GetField('PCF_CALCFRACTION')+';'+GetControlText('VALIDITE');
        AGLLanceFiche('PAY','CALCFRACDETAIL','','','MODIF;'+St);
        Combo := THValComboBox(GeTcontrol('VALIDITE'));
        Q := OpenSQL('SELECT PCD_DATEDEBUT FROM CALCFRACTIONDETAIL '+
        'WHERE PCD_ORDRE=1 AND PCD_CALCFRACTION="'+GetField('PCF_CALCFRACTION')+'" ORDER BY PCD_DATEDEBUT DESC',True);
        TobDetail := Tob.Create('Les fractions',Nil,-1);
        TobDetail.LoadDetailDB('CALCFRACTIONDETAIL','','',Q,False);
        Ferme(Q);
        Combo.Items.Clear;
        Combo.Values.Clear;
        ValeurCombo := '';
        For i := 0 to TobDetail.Detail.Count -1 do
        begin
                if i = 0 then ValeurCombo := DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT'));
                Combo.Items.Add('à partir du '+DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT')));
                Combo.Values.Add(DateToStr(TobDetail.Detail[i].Getvalue('PCD_DATEDEBUT')));
        end;
        Combo.Value := ValeurCombo;
        AfficherGrilleRetenue;
end;

Initialization
  registerclasses ( [ TOM_CALCFRACTION ] ) ;
end.

