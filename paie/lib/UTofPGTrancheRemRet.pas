{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGTRANCHEREMRET ()
Mots clefs ... : TOF;PGTRANCHEREMRET
*****************************************************************
PT1  : 16/01/2004 JL V_50 FQ 11000 Modif bouton calcul
PT2  : 01/06/2004 JL V_50 FQ 11321 CORRECTION CALCUL RMI : ajour ordre de tri pour reprendre le dernier
PT3  : 05/08/2004 JL V_50 FQ 11469 Redimensionnement grille sur OnArgument
PT4  : 12/08/2004 JL V_50 FQ 11148 Gestion acomptes
PT5  : 29/10/2004 JL V_60 FQ 11743 Conversion de type variant incorrect en simulation salarié
                             + correction faute  d'orthographe le 14/12/2004   
PT6  : 06/02/2007 FC V_80 Mise en place de la gestion des habilitations/populations
PT7  : 31/07/2007 JL V_80 FQ 13982 Gestion valeur 90 et 99 pour champ personne a charge
}
Unit UTofPGTrancheRemRet ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,Spin,HTB97,HSysMenu,P5DEF,EntPaie,PGOutils2 ;

Type
  TOF_PGTRANCHEREMRET = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure AfficherPeriode (Sender : TObject);
    procedure CalculerPartieSaisissable;
    procedure CalculerPeriodeSalarie;
    procedure ChangeMethodeCalc (Sender : TObject);
    procedure BCalcClick (Sender : TObject);
    procedure ExitEdit(Sender : TObject);
  end ;

Implementation

procedure TOF_PGTRANCHEREMRET.OnArgument (S : String ) ;
var Combo : THValComboBox;
    Q : TQuery;
    TobPer : Tob;
    i : Integer;
    Valeur,StDate,Libelle,Salarie,CodeElement,StDateDeb,StDateFin : String;
    BCalculer : TToolBarButton97;
    DatedebPaie,DateFinPaie : TDateTime;
    RMI : Double;
    HMTrad: THSystemMenu;
    Edit : THEdit;
begin
  Inherited ;
        Salarie := ReadTokenPipe(S,';');
        Edit := THEdit(GetControl('SALARIE'));
        Edit.OnExit := ExitEdit;
        If Salarie <> '' then
        begin
                StDateDeb := ReadTokenPipe(S,';');
                StDateFin := ReadTokenPipe(S,';');
                SetControlText('SALARIE',Salarie);
                SetControlENabled('SALARIE',False);
                If IsValidDate (StDateDeb) then
                begin
                        SetControlText('DATEDEBSAL',StDateDeb);
                        SetControlEnabled('DATEDEBSAL',False);
                end;
                If IsValidDate (StDateFin) then
                begin
                        SetControlText('DATEFINSAL',StDateFin);
                        SetControlEnabled('DATEFINSAL',False);
                end;
        end;

        Combo := THValComboBox(GetControl('PERIODEREM'));
        If combo <> Nil then Combo.OnChange := AfficherPeriode;
        Q := OpenSQL('SELECT PCF_CALCFRACTION FROM CALCFRACTION',True);
        If Not Q.Eof then SetControlText('METHODECALC',Q.FindField('PCF_CALCFRACTION').AsString);
        Ferme(Q);
        Q := OpenSQL('SELECT PCD_DATEDEBUT,PCD_DATEFIN,PCF_CODEELT FROM CALCFRACTIONDETAIL '+
        'LEFT JOIN CALCFRACTION ON PCF_CALCFRACTION=PCD_CALCFRACTION '+
        'WHERE PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" AND '+
        'PCD_ORDRE=1'+
        ' ORDER BY PCD_DATEDEBUT DESC',True);
        TobPer := Tob.Create('Les périodes',Nil,-1);
        TobPer.LoadDetailDB('Les périodes','','',Q,False);
        Ferme(Q);
        For i := 0 to TobPer.Detail.Count-1 do
        begin
                        StDate := DateToStr(TobPer.Detail[i].GetValue('PCD_DATEDEBUT'));
                        If TobPer.Detail[i].GetValue('PCD_DATEFIN') = IDate1900 then
                        Libelle := 'à partir du '+StDate
                        Else Libelle := 'Du '+ StDate+' au '+DateToStr(TobPer.Detail[i].GetValue('PCD_DATEFIN'));
                        Combo.Items.Add(Libelle);
                        Combo.Values.Add(StDate);
                        If i = 0 then Valeur := StDate;
                        CodeElement := TobPer.Detail[i].GetValue('PCF_CODEELT');
        end;
        SetControlText('PERIODEREM',Valeur);
        AfficherPeriode(Nil);
        TobPer.free;
        BCalculer := TToolBarButton97(GetControl('BCALCULER'));
        If BCalculer <> Nil Then BCalculer.OnClick := BCalcClick;  //PT1
        DatedebPaie := DebutDeMois(PlusMois(Date,-1));
        DateFinPaie := FinDeMois(DatedebPaie);
        SetControlText('DATEDEBSAL',DateToStr(DateDebPaie));
        SetControlText('DATEFINSAL',DateToStr(DateFinPaie));
        SetControlText('DATEPAIE',DateToStr(Date));
        SetControlText('TYPERET','FSA');
        Q := OpenSQL('SELECT PEL_MONTANTEURO FROM ELTNATIONAUX WHERE PEL_CODEELT="'+CodeElement+'" AND '+
        'PEL_PREDEFINI="CEG" AND PEL_DATEVALIDITE<="'+UsDateTime(Date)+'" ORDER BY PEL_DATEVALIDITE DESC',true); // PT2
        If Not Q.Eof then RMI := Q.FindField('PEL_MONTANTEURO').AsFloat
        Else RMI := 0;
        Ferme(Q);
        SetControlText('THRMI',FloatToStr(RMI));
        Combo := THValComboBox(GetControl('METHODECALC'));
        If Combo <> Nil then Combo.OnChange := ChangeMethodeCalc;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GSALARIE'))) ;    //PT3
        HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLEREM'))) ;
end ;

procedure TOF_PGTRANCHEREMRET.AfficherPeriode (Sender : TObject);
var Q : TQuery;
    TobRem,TobGrille,T : Tob;
    StDate,StInf,StSup : String;
    GrilleRem : THGrid;
    i : Integer;
    MontantInf,MontantSup : Double;
    Numerateur,Denominateur : Integer;
    HMTrad: THSystemMenu;
begin
        GrilleRem := THGrid(GetControl('GRILLEREM'));
        For i := 1 to GrilleRem.RowCount -1 do
        begin
                GrilleRem.Rows[i].Clear;
        end;
        GrilleRem.RowCount := 2;
        GrilleRem.ColAligns[0]:=taRightJustify;
        GrilleRem.ColAligns[1]:=taRightJustify;
        GrilleRem.ColAligns[2]:=taCenter;
        GrilleRem.ColAligns[3]:=taRightJustify;
        StDate := GetControlText('PERIODEREM');
        Q := OpenSQL('SELECT PCD_MONTANTINF,PCD_MONTANTSUP,PCD_NUMERATEUR,PCD_DENOMINATEUR FROM CALCFRACTIONDETAIL'+
        ' WHERE PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" AND '+
        'PCD_DATEDEBUT="'+UsDateTime(StrToDate(StDate))+'" AND PCD_ORDRE>0 ORDER BY PCD_ORDRE',True);
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
        ' WHERE PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" AND '+
        'PCD_DATEDEBUT="'+UsDateTime(StrToDate(StDate))+'" AND PCD_ORDRE=1',True);
        If Not Q.Eof then SetControlText('PERSACHARGE',FloatToStr(Q.FindField('PCD_AUGACHARGE').AsFloat));
        Ferme(Q);
        HMTrad.ResizeGridColumns(GrilleRem) ;
end;

procedure TOF_PGTRANCHEREMRET.BCalcClick (Sender : TObject);            //PT1
var Page : TPageControl;
begin
        Page := TPageControl(GetControl('PAGES'));
        If Page.ActivePage.Name = 'PGENERAL' then CalculerPartieSaisissable
        else CalculerPeriodeSalarie;
end;
procedure TOF_PGTRANCHEREMRET.CalculerPartieSaisissable ;
var SalaireNet,SalaireInf,SalaireSup,AugACharge,SalaireCalc,RMI : Double;
    Total,TotalTranche : Double;
    TobTrancheRem,TobRem,TobGrille,T : Tob;
    Q : TQuery;
    StDate,StInf,StSup,TypeRetenue,Partie : String;
    i,NbAcharge,Numerateur,Denominateur : Integer;
    SpEdit : TSpinEdit;
    GrilleRem : THGrid;
    HMTrad: THSystemMenu;
begin
        If not IsNumeric(GetControlText('SALAIRENET')) then
        begin
                PGIBox('Le salaire net doit être numérique',Ecran.Caption);
                Exit;
        end;
        SalaireNet := StrToFloat(GetControlText('SALAIRENET'));
        If SalaireNet <= RMI then
                begin
                        PGIBox('Le salaire est non saisissable en dessous de la valeur du RMI');
                        SetControlText('TOTALSAISIE','0');
                        Exit;
                end;
        SpEdit := TSpinEdit(Getcontrol('NBACHARGE'));
        If SpEdit <> Nil Then NbACharge := SpEdit.Value
        else NbACharge := 0;
        GrilleRem := THGrid(GetControl('GRILLEREM'));
        For i := 1 to GrilleRem.RowCount -1 do
        begin
                GrilleRem.Rows[i].Clear;
        end;
        GrilleRem.RowCount := 2;
        GrilleRem.ColAligns[0]:=taRightJustify;
        GrilleRem.ColAligns[1]:=taRightJustify;
        GrilleRem.ColAligns[2]:=taCenter;
        GrilleRem.ColAligns[3]:=taRightJustify;
        TypeRetenue := GetControlText('TYPERET');
        Q := OpenSQL('SELECT PTR_TYPEFRACTION FROM TYPERETENUE WHERE PTR_RETENUESAL="'+TypeRetenue+'"',True);
        If not Q.Eof then Partie := Q.FindField('PTR_TYPEFRACTION').AsString;
        Ferme(Q);
//        If Partie = 'FAI' then
//        begin
//                PGIBox('Pour ce type de retenue tout le salaire est saisissable#13#10soit '+FloatToStr(SalaireNet),Ecran.Caption);
//                SetControlText('TOTALSAISIE',FloatToStr(SalaireNet));
//        end;
        If (Partie = 'FRI') or (Partie = 'FAI')  then
        begin
                RMI := StrToFloat(GetControlText('THRMI'));
                PGIBox('Pour ce type de retenue tout le salaire est saisissable sauf la part égale au RMI'+
                '#13#10soit '+FloatToStr(SalaireNet - RMI),Ecran.Caption);
                SetControlText('TOTALSAISIE',FloatToStr(SalaireNet - RMI));
        end;
        If Partie = 'F10' then
        begin
                PGIBox('Pour ce type de retenue, 10% du salaire net est saisissable'+
                '#13#10soit '+FloatToStr(SalaireNet*10/100),Ecran.Caption);
                SetControlText('TOTALSAISIE',FloatToStr(SalaireNet*10/100));
        end;
        If Partie = 'FSA' then
        begin
                RMI := StrToFloat(GetControlText('THRMI'));
                StDate := GetControlText('PERIODEREM');
                Q := OpenSQL('SELECT PCD_MONTANTINF,PCD_MONTANTSUP,PCD_NUMERATEUR,PCD_DENOMINATEUR,PCD_AUGACHARGE FROM CALCFRACTIONDETAIL'+
                ' WHERE PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" AND '+
                'PCD_DATEDEBUT="'+UsDateTime(StrToDate(StDate))+'" AND PCD_ORDRE>0 ORDER BY PCD_ORDRE',True);
                TobRem := Tob.Create('Les rémunérations',Nil,-1);
                TobRem.LoadDetailDB('Les rémunérations','','',Q,False);
                Ferme(Q);
                TobGrille := Tob.Create('Remplir Grille',Nil,-1);
                For i:= 0 to TobRem.detail.Count - 1 do
                begin
                        SalaireInf := TobRem.Detail[i].GetValue('PCD_MONTANTINF');
                        SalaireSup := TobRem.Detail[i].getValue('PCD_MONTANTSUP');
                        Numerateur := TobRem.Detail[i].GetValue('PCD_NUMERATEUR');
                        Denominateur := TobRem.Detail[i].GetValue('PCD_DENOMINATEUR');
                        AugACharge := TobRem.Detail[i].GetValue('PCD_AUGACHARGE');
                        SalaireInf := SalaireInf + (AugACharge * NbACharge);
                        If SalaireSup > 0 then SalaireSup := SalaireSup + (AugACharge * NbACharge);
                        StInf := FloatToStr(SalaireInf);
                        If SalaireSup > 0 then StSup := FloatToStr(SalaireSup)
                        Else StSup := '/';
                        T := Tob.Create('Une fille',TobGrille,-1);
                        T.AddChampSupValeur('INFERIEUR',StInf,False);
                        T.AddChampSupValeur('SUPERIEUR',StSup,False);
                        T.AddChampSupValeur('FRACTION',IntToStr(Numerateur)+'/'+IntToStr(Denominateur),False);
                end;
                TobRem.Free;
                TobGrille.PutGridDetail(GrilleRem,False,True,'',False);
                TobGrille.Free;
                Q := OpenSQL('SELECT PCD_AUGACHARGE FROM CALCFRACTIONDETAIL'+
                ' WHERE PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" AND '+
                'PCD_DATEDEBUT="'+UsDateTime(StrToDate(StDate))+'" AND PCD_ORDRE=1',True);
                If Not Q.Eof then SetControlText('PERSACHARGE',FloatToStr(Q.FindField('PCD_AUGACHARGE').AsFloat));
                Ferme(Q);
                GrilleRem := THGrid(GetControl('GRILLEREM'));
                GrilleRem.ColCount := 4;
                StDate := GetControlText('PERIODEREM');
                Q := OpenSQL('SELECT PCD_MONTANTINF,PCD_MONTANTSUP,PCD_AUGACHARGE,PCD_DENOMINATEUR,PCD_NUMERATEUR FROM CALCFRACTIONDETAIL'+
                ' WHERE PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" AND '+
                'PCD_DATEDEBUT="'+UsDateTime(StrToDate(StDate))+'" AND PCD_ORDRE>0 ORDER BY PCD_ORDRE',True);
                TobTrancheRem := Tob.Create('Les rémunérations',Nil,-1);
                TobTrancheRem.LoadDetailDB('Les rémunérations','','',Q,False);
                Ferme(Q);
                Total := 0;
                For i := 0 to TobTrancheRem.Detail.Count - 1 do
                begin
                        SalaireInf := TobTrancheRem.Detail[i].GetValue('PCD_MONTANTINF');
                        SalaireSup := TobTrancheRem.Detail[i].GetValue('PCD_MONTANTSUP');
                        AugACharge := TobTrancheRem.Detail[i].GetValue('PCD_AUGACHARGE');
                        Numerateur := TobTrancheRem.Detail[i].GetValue('PCD_NUMERATEUR');
                        Denominateur := TobTrancheRem.Detail[i].GetValue('PCD_DENOMINATEUR');
                        If SalaireInf > 0 then SalaireInf := (AugACharge * NbACharge) + SalaireInf;
                        If SalaireSup > 0 then SalaireSup := (AugACharge * NbACharge) + SalaireSup;
                        If SalaireNet > SalaireInf then
                        begin
                                If (SalaireNet > SalaireSup) and (SalaireSup > 0) then SalaireCalc := SalaireSup - SalaireInf
                                Else SalaireCalc := SalaireNet - SalaireInf;
                                TotalTranche := arrondi((SalaireCalc * Numerateur)/Denominateur,2);
                                If SalaireNet - (Total + TotalTranche) < RMI then TotalTranche := SalaireNet - Total - RMI;
                                Total := Total + TotalTranche;
                                GrilleRem.CellValues[3,i+1] := FloatToStr(TotalTranche);
                        end
                        Else GrilleRem.CellValues[3,i+1] := IntToStr(0);
                end;
                TobTrancheRem.Free;
                SetControlText('TOTALSAISIE',FloatToStr(arrondi(Total,2)));
                HMTrad.ResizeGridColumns(GrilleRem) ;
        end;
end;

procedure TOF_PGTRANCHEREMRET.CalculerPeriodeSalarie;
var GrilleSal : THGrid;
    TobRetenue,TobPaie : Tob;
    Q : TQuery;
    i,NbACharge,NbPaie : Integer;
    Salarie : String;
    HMTrad: THSystemMenu;
    SalaireNet,MtAcpt : Double;
    Acompte : String ;
   TobAcSaisieArret : Tob;
    Habilitation:String;//PT6
    Stop:Boolean;//PT6
begin
        //DEB PT6
        Habilitation := '';
        if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
          Habilitation := ' AND PPU_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
        //FIN PT6

        Salarie := GetControlText('SALARIE');
        if GetCheckBoxState('CBPAIE') = CbChecked then
        begin
                Q := OpenSQL('SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_CBRUT FROM PAIEENCOURS WHERE PPU_SALARIE="'+Salarie+'"'+
                ' AND PPU_DATEDEBUT>="'+UsDateTime(StrToDate(Getcontroltext('DATEPAIE')))+'" ' + Habilitation,True); //PT6
                TobPaie := Tob.Create('Les Paies',Nil,-1);
                TobPaie.LoadDetailDB('Les Paies','','',Q,False);
                Ferme(Q);
                SalaireNet := 0;
                NbPaie := 0;
                If TobPaie.Detail.Count = 0 then
                begin
                        PGIBox('Aucun bulletin de paie trouvé',Ecran.Caption);
                        TobPaie.Free;
                        Exit;
                end;
                For i := 0 to TobPaie.Detail.Count - 1 do
                begin
                        NbPaie := NbPaie + 1;
                        SalaireNet := SalaireNet + TobPaie.Detail[i].GetValue('PPU_CBRUT');
                end;
                If NbPaie > 0 then SalaireNet := SalaireNet / NbPaie;
                TobPaie.Free;
        end
        else
        begin
                SalaireNet := 0;
                Q := OpenSQL('SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_CBRUT FROM PAIEENCOURS WHERE PPU_SALARIE="'+Salarie+'"'+
                ' AND PPU_DATEDEBUT="'+UsDateTime(StrToDate(Getcontroltext('DATEDEBSAL')))+'" AND PPU_DATEFIN="'+UsDateTime(StrToDate(Getcontroltext('DATEFINSAL')))+'" '+ Habilitation,True);   //PT6
                If Not Q.Eof then SalaireNet := Q.FindField('PPU_CBRUT').AsFloat
                else
                begin
                        PGIBox('Aucun bulletin de paie trouvé, prise en compte du salaire net saisi.',Ecran.Caption);   //PT15
//                        Exit;
                end;
                Ferme(Q);
                // Acomptes PT4
                If VH_Paie.PGGESTACPT = true then
                begin
                        MtAcpt :=0;
                        Q := OpenSQL('SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_SAISIEARRETAC="X"',True);
                        TobAcSaisieArret := Tob.Create('Les rubriques acomptes',Nil,-1);
                        TobAcSaisieArret.LoadDetailDB('Les acomptes','','',Q,False);
                        Ferme(Q);
                        For i := 0 to TobAcSaisieArret.Detail.Count - 1 do
                        begin
                          Acompte := TobAcSaisieArret.Detail[i].GetValue('PRM_RUBRIQUE');

                          //DEB PT6
                          Habilitation := '';
                          if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
                            Habilitation := ' AND PHB_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
                          //FIN PT6

                          Q := OpenSQL('SELECT SUM(PHB_MTREM) MTREM FROM HISTOBULLETIN '+
                          'WHERE PHB_NATURERUB="AAA" AND PHB_SALARIE="'+Salarie+'" AND PHB_RUBRIQUE="'+Acompte+'"'+
                          ' AND PHB_DATEDEBUT="'+UsDateTime(StrToDate(Getcontroltext('DATEDEBSAL')))+'" '+
                          'AND PHB_DATEFIN="'+UsDateTime(StrToDate(Getcontroltext('DATEFINSAL')))+'" ' + Habilitation,True);   //PT6
                          If Not Q.Eof then MtAcpt := Q.FindField('MTREM').AsInteger
                          else MtAcpt := 0;
                          Ferme(Q);
                        end;
                TobAcSaisieArret.Free;
                end
                else MtAcpt :=0;
                SalaireNet := SalaireNet + MtAcpt;
        end;
        If SalaireNet > 0 then SetControlText('SALAIRENETMOY',FloatToStr(Salairenet))
        else
        begin           //DEBUT PT5
                If Not IsNumeric(GetControlText('SALAIRENETMOY')) then
                begin
                        PGIBox('Le salaire net doit être numérique',Ecran.Caption);
                        Exit;
                end;
                 SalaireNet := StrToFloat(GetControlText('SALAIRENETMOY'));
        end;            //FIN PT5
        GrilleSal := THGrid(GetControl('GSALARIE'));
        For i := 1 to GrilleSal.RowCount -1 do
        begin
                GrilleSal.Rows[i].Clear;
        end;
        GrilleSal.ColAligns[0] := taLeftJustify;
        GrilleSal.ColAligns[1] := taLeftJustify;
        GrilleSal.ColAligns[2] := taRightJustify;
        GrilleSal.ColAligns[3] := taRightJustify;
        //DEB PT6
        Habilitation := '';
        if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
          Habilitation := ' AND ' + MonHabilitation.LeSQL;
        //FIN PT6
        Q := OpenSQL('SELECT PSA_PERSACHARGE FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"' + Habilitation,True); //PT6
        Stop := False; //PT6
        if not Q.Eof then
        begin
          NbACharge := Q.FindField('PSA_PERSACHARGE').AsInteger;
          If (NbAcharge=90) or (NbAcharge=99) then NbAcharge :=0;//PT7
        end
        else
        begin
          NbACharge := 0;
          Stop := True; //PT6
        end;
        Ferme(Q);
        if Not Stop then  //PT6 Inutile de faire la suite si le salarié n'est pas accessible
        begin
          TobRetenue := PGCalcSaisieArret(Salarie,StrToDate(GetControltext('DATEDEBSAL')),StrToDate(GetControltext('DATEFINSAL')),NbACharge,StrToFloat(GeTcontrolText('SALAIRENETMOY')),0);
          if TobRetenue <> Nil then
          begin
        //          PGMajHistoRetenue(TobRetenue,StrToDate(GetControltext('DATEFINSAL')));
                  TobRetenue.PutGridDetail(GrilleSal,False,True,'PTR_LIBELLE;PRE_LIBELLE;PRE_MONTANTMENS;MONTANTBUL',False);
                  TobRetenue.Free;
          end;
        end;
        HMTrad.ResizeGridColumns(GrilleSal) ;
end;

procedure TOF_PGTRANCHEREMRET.ChangeMethodeCalc (Sender : TObject);
var Q : TQuery;
    TobDetail : Tob;
    combo : THValComboBox;
    i : Integer;
    ValeurCombo : String;
begin
        Combo := THValComboBox(GeTcontrol('PERIODEREM'));
        Q := OpenSQL('SELECT PCD_DATEDEBUT FROM CALCFRACTIONDETAIL '+
        'WHERE PCD_ORDRE=1 AND PCD_CALCFRACTION="'+GetControlText('METHODECALC')+'" ORDER BY PCD_DATEDEBUT DESC',True);
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
end;

procedure TOF_PGTRANCHEREMRET.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

Initialization
  registerclasses ( [ TOF_PGTRANCHEREMRET ] ) ;
end.

