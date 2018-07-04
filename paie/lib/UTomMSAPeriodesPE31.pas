{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : MSAPERIODESPE31 (MSAPERIODESPE31)
Mots clefs ... : TOM;MSAPERIODESPE31
*****************************************************************
PT1 19/05/2004 JL V_60 FQ 12294 Unité de gestion sur 2 caractères
PT2 04/04/2006 JL V_65 Gestion UG MSA
PT3 11/05/2007 JL V_72 FQ 13888 matricule salarié complété automatiquement
PT4 12/11/2007 FC V_80 FQ 14892 Evolution cahier des charges Octobre 2007
PT5 02/01/2008 FC V_81 FQ 15091 Amélio MSA
PT7 12/03/2007 FC V_80 FQ 15303 2 rubriques d'heures TEPA sur le taux de majoration 25
PT8 31/03/2008 FC V_80 FQ 15333 correction cahier des charges
PT9 20/05/2008 FC V_90 FQ 15395 Suppression du controle sur le type de rémunération
}
Unit UTomMSAPeriodesPE31 ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,
{$else}
     eFiche,eFichList,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,LookUp,EntPaie,PGOutils2 ;

Type
  TOM_MSAPERIODESPE31 = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnAfterDeleteRecord             ; override ;  //PT7
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    GRem : THGrid;
    MemSalarie:String; //PT7
    MemDateDebut,MemDateFin : TDateTime;  //PT7
    procedure RemplirGrilleRemuneration ;
    procedure GrilleElipsisClick(Sender: TObject);
    procedure GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure ExitEdit(Sender: TObject);   //PT3
    procedure cbComboChange(Sender: TObject); //PT4
    procedure cbValChange(Sender: TObject); //PT4
    end ;

Implementation

procedure TOM_MSAPERIODESPE31.OnNewRecord ;
begin
  Inherited ;
        SetField('PE3_ORDRE',0);
        SetField('PE3_DATEDEBUT',DebutDeMois(Date));
        SetField('PE3_DATEFIN',FinDeMois(Date));
end ;

procedure TOM_MSAPERIODESPE31.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_MSAPERIODESPE31.OnUpdateRecord ;
var i : Integer;
    MessError : String;
    TotalHeures : Double;
    TypeRem : String;
    Taux:String;//PT5
begin
  Inherited ;
        MessError := '';
        If GetField('PE3_SALARIE') = '' then MessError := MessError+'#13#10 - Le matricule du salarié';
        If GetField('PE3_ETABLISSEMENT') = '' then MessError := MessError+'#13#10 - L''établissement du salarié';
        If GetField('PE3_MSAACTIVITE') = '' then MessError := MessError+'#13#10 - L''activité du salarié';
        If GetField('PE3_DATEDEBUT') = IDate1900 then MessError := MessError+'#13#10 - La date de début';
        If GetField('PE3_DATEFIN') = IDate1900 then MessError := MessError+'#13#10 - La date de fin';
        If GRem.CellValues[0,1] = '' then MessError := MessError+'#13#10 - Le premier élément de calcul';
        If GRem.CellValues[2,1] = '' then MessError := MessError+'#13#10 - Le montant du premier élément de calcul';
        If MessError <> '' then
        begin
                LastError := 1;
                PGIBox('La (ou les) information(s) suivante(s) sont obligatoire(s) :'+MessError,Ecran.Caption);
                exit;
        end;
        For i := 1 to GRem.RowCount - 1 do
        begin
                If GRem.CellValues[0,i]<> '' then
                begin
                       If Not IsNumeric(GRem.CellValues[2,i]) then MessError := '#13#10- le montant de la rémunération '+GRem.CellValues[1,i]+' n''est pas numérique';
                end;
        end;
        If MessError <> '' then
        begin
                LastError := 1;
                PGIBox('Les informations suivantes sont erronées :'+MessError,Ecran.Caption);
                exit;
        end;
        
        //DEB PT4
        //Vérifier le type de rémunération
        TotalHeures := 0;
        for i := 1 to 4 do
        begin
          if (GetControlText('TYPEREM' + IntToStr(i)) <> '') then
          begin
            if TotalHeures = 0 then   //PT7  //PT8
              TypeRem := '30'
            else
              TypeRem := '40';
            //DEB PT5
            if i = 1 then Taux := '25%';
            if i = 2 then Taux := '50%';
            if i = 3 then Taux := 'suppl 01';
            if i = 4 then Taux := 'suppl 02';
            //FIN PT5
//PT9            if TypeRem <> GetControlText('TYPEREM' + IntToStr(i)) then
//PT9              MessError := MessError + '#13#10- le type de rémunération correspondant au taux de majoration ' + Taux + ' devrait être égal à ' + TypeRem; //PT5
            if GetControlText('TAUXMAJO' + IntToStr(i)) = '' then
              MessError := MessError + '#13#10- le taux de majoration ' + Taux + ' est obligatoire';  //PT5
            if GetControlText('NBHEURES' + IntToStr(i)) = '' then
              MessError := MessError + '#13#10- le nombre d''heures correspondant au taux de majoration ' + Taux + ' est obligatoire' //PT5
            else
              TotalHeures := TotalHeures + StrToFloat(GetControlText('NBHEURES' + IntToStr(i)));
            if GetControlText('MONTANTREM' + IntToStr(i)) = '' then
              MessError := MessError + '#13#10- le montant correspondant au taux de majoration ' + Taux + ' est obligatoire'; //PT5
          end;
        end;
        If MessError <> '' then
        begin
                LastError := 1;
                PGIBox('Les informations suivantes sont erronées :'+MessError,Ecran.Caption);
                exit;
        end;
        //FIN PT4

        For i := 1 to GRem.RowCount - 1 do
        begin
                If GRem.CellValues[0,i] <> '' then
                begin
                        SetField('PE3_ELTCALCMSA'+ IntToStr(i),GRem.CellValues[0,i]);
                        SetField('PE3_MONTANTELT'+ IntToStr(i),Arrondi(StrToFloat(GRem.CellValues[2,i]),0));
                end
                else
                begin
                        SetField('PE3_ELTCALCMSA'+ IntToStr(i),'');
                        SetField('PE3_MONTANTELT'+ IntToStr(i),0);
                end;
        end;
end ;

procedure TOM_MSAPERIODESPE31.OnAfterUpdateRecord ;
var i : Integer;
    TobMSA : Tob;
begin
  Inherited ;
        If GRem.CellValues[0,6]<> '' then
        begin
                TobMSA := Tob.Create('MSAPERIODESPE31',Nil,-1);
                TobMSA.PutValue('PE3_ETABLISSEMENT',GetField('PE3_ETABLISSEMENT'));
                TobMSA.PutValue('PE3_SALARIE',GetField('PE3_SALARIE'));
                TobMSA.PutValue('PE3_ORDRE',1);
                TobMSA.PutValue('PE3_DATEDEBUT',GetField('PE3_DATEDEBUT'));
                TobMSA.PutValue('PE3_DATEFIN',GetField('PE3_DATEFIN'));
                TobMSA.PutValue('PE3_NOM',GetField('PE3_NOM'));
                TobMSA.PutValue('PE3_PRENOM',GetField('PE3_PRENOM'));
                TobMSA.PutValue('PE3_DATENAISSANCE',GetField('PE3_DATENAISSANCE'));
                TobMSA.PutValue('PE3_NUMEROSS',GetField('PE3_NUMEROSS'));
                TobMSA.PutValue('PE3_NBHEURES',GetField('PE3_NBHEURES'));
                TobMSA.PutValue('PE3_NBJOURS',GetField('PE3_NBJOURS'));
                TobMSA.PutValue('PE3_MSAACTIVITE',GetField('PE3_MSAACTIVITE'));
                TobMSA.PutValue('PE3_UGMSA',GetField('PE3_UGMSA'));
                TobMSA.PutValue('PE3_TOPPLAF',GetField('PE3_TOPPLAF'));
                TobMSA.PutValue('PE3_PEXOMSA',GetField('PE3_PEXOMSA'));
                For i := 6 to 10 do
                begin
                        TobMSA.PutValue('PE3_ELTCALCMSA'+IntToStr(i-5),GRem.CellValues[0,i]);
                        If IsNumeric(GRem.CellValues[2,i]) then TobMSA.PutValue('PE3_MONTANTELT'+IntToStr(i-5),StrToInt(GRem.CellValues[2,i]))
                        else TobMSA.PutValue('PE3_MONTANTELT'+IntToStr(i-5),0);

                end;
                TobMSA.InsertOrUpdateDB();
                TobMSA.Free;
        end
        else ExecuteSQL('DELETE FROM MSAPERIODESPE31 WHERE PE3_SALARIE="'+GetField('PE3_SALARIE')+'" AND '+
        'PE3_DATEDEBUT="'+UsDateTime(GetField('PE3_DATEDEBUT'))+'" AND PE3_DATEFIN="'+UsDateTime(GetField('PE3_DATEFIN'))+'" AND PE3_ORDRE>0');

      //DEB PT4
      TobMSA := Tob.Create('MSAPERIODESPE31',Nil,-1);
      TobMSA.PutValue('PE3_ETABLISSEMENT',GetField('PE3_ETABLISSEMENT'));
      TobMSA.PutValue('PE3_SALARIE',GetField('PE3_SALARIE'));
      TobMSA.PutValue('PE3_DATEDEBUT',GetField('PE3_DATEDEBUT'));
      TobMSA.PutValue('PE3_DATEFIN',GetField('PE3_DATEFIN'));
      TobMSA.PutValue('PE3_NOM',GetField('PE3_NOM'));
      TobMSA.PutValue('PE3_PRENOM',GetField('PE3_PRENOM'));
      TobMSA.PutValue('PE3_DATENAISSANCE',GetField('PE3_DATENAISSANCE'));
      TobMSA.PutValue('PE3_NUMEROSS',GetField('PE3_NUMEROSS'));
      TobMSA.PutValue('PE3_NBHEURES',0);
      TobMSA.PutValue('PE3_NBJOURS',0);
      TobMSA.PutValue('PE3_MSAACTIVITE',GetField('PE3_MSAACTIVITE'));
      TobMSA.PutValue('PE3_UGMSA',GetField('PE3_UGMSA'));
      TobMSA.PutValue('PE3_TOPPLAF','-');
      TobMSA.PutValue('PE3_PEXOMSA','');

      for i := 0 to 3 do
      begin
        if GetControlText('TYPEREM' + IntToStr(i + 1)) <> '' then
        begin
          TobMSA.PutValue('PE3_ORDRE',i + 100);
          TobMSA.PutValue('PE3_MONTANTELT1',GetControlText('TYPEREM' + IntToStr(i + 1)));
          TobMSA.PutValue('PE3_MONTANTELT2',GetControlText('TAUXMAJO' + IntToStr(i + 1)));
          TobMSA.PutValue('PE3_MONTANTELT3',StrToFloat(GetControlText('NBHEURES' + IntToStr(i + 1))) * 100);
          TobMSA.PutValue('PE3_MONTANTELT4',GetControlText('MONTANTREM' + IntToStr(i + 1)));
          TobMSA.InsertOrUpdateDB();
        end
        else
          ExecuteSQL('DELETE FROM MSAPERIODESPE31 WHERE PE3_SALARIE="'+GetField('PE3_SALARIE')+'" AND '+
            ' PE3_DATEDEBUT="'+UsDateTime(GetField('PE3_DATEDEBUT'))+'"' +
            ' AND PE3_DATEFIN="'+UsDateTime(GetField('PE3_DATEFIN'))+'" AND PE3_ORDRE=' + IntToStr(i + 100));
      end;
      TobMSA.Free;
      //FIN PT4
end ;

procedure TOM_MSAPERIODESPE31.OnLoadRecord ;
begin
  Inherited ;
  RemplirGrilleRemuneration;
  TFFiche(Ecran).Caption := 'Périodes MSA, salarié : '+getField('PE3_SALARIE')+' '+GetField('PE3_NOM')+' '+GetField('PE3_PRENOM');
  UpdateCaption(TFFiche(Ecran));

  //DEB PT7
  MemSalarie := getField('PE3_SALARIE');
  MemDateDebut := getField('PE3_DATEDEBUT');
  MemDateFin := getField('PE3_DATEFIN');
  //FIN PT7
end ;

procedure TOM_MSAPERIODESPE31.OnChangeField ( F: TField ) ;
var Q : TQuery;
begin
  Inherited ;
        If (F.FieldName = 'PE3_SALARIE') and ((DS.State in [dsInsert])) then
        begin
          If (TFFiche(Ecran).FTypeAction = TaCreat) and (GetField('PE3_SALARIE') <> '') Then ExitEdit(THEdit(GetControl('PE3_SALARIE')));//PT3
                If GetField('PE3_SALARIE') <> '' then
                begin
                        Q := OpenSQL('SELECT PSE_MSAUNITEGES,PSA_SALARIE,PSA_LIBELLE,PSA_NUMEROSS,PSA_ETABLISSEMENT,PSA_PRENOM,PSA_DATENAISSANCE,PSE_MSAACTIVITE FROM SALARIES'+
                        ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE'+
                        ' WHERE PSA_SALARIE="'+GetField('PE3_SALARIE')+'"',True);
                        If Not Q.Eof then
                        begin
                                SetField('PE3_ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
                                SetField('PE3_NOM',Q.FindField('PSA_LIBELLE').AsString);
                                SetField('PE3_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);
                                SetField('PE3_PRENOM',Q.FindField('PSA_PRENOM').AsString);
                                SetField('PE3_DATENAISSANCE',Q.FindField('PSA_DATENAISSANCE').AsDateTime);
                                SetField('PE3_MSAACTIVITE',Q.FindField('PSE_MSAACTIVITE').AsString);
                                SetField('PE3_UGMSA',Q.FindField('PSE_MSAUNITEGES').AsString);    //PT2
                        end;
                        Ferme(Q);
                        TFFiche(Ecran).Caption := 'Saisie Période MSA : salarié '+getField('PE3_SALARIE')+' '+GetField('PE3_NOM')+' '+GetField('PE3_PRENOM');
                        UpdateCaption(TFFiche(Ecran));
                end;
        end;

  If (F.FieldName = 'TYPEREM1') or (F.FieldName = 'TYPEREM2') or (F.FieldName = 'TYPEREM3') or (F.FieldName = 'TYPEREM4')
   or(F.FieldName = 'TAUXMAJO1') or (F.FieldName = 'TAUXMAJO2') or (F.FieldName = 'TAUXMAJO3') or (F.FieldName = 'TAUXMAJO4')
   or(F.FieldName = 'NBHEURES1') or (F.FieldName = 'NBHEURES2') or (F.FieldName = 'NBHEURES3') or (F.FieldName = 'NBHEURES4')
   or(F.FieldName = 'MONTANTREM1') or (F.FieldName = 'MONTANTREM2') or (F.FieldName = 'MONTANTREM3') or (F.FieldName = 'MONTANTREM4') then
     SetField('PE3_SALARIE',GetField('PE3_SALARIE'));

end ;

procedure TOM_MSAPERIODESPE31.OnArgument ( S: String ) ;
var PInv : TTabSheet;
  i : integer;
begin
  Inherited ;
        PInv := TTabSheet(GetControl('PINVISIBLE'));
        PInv.TabVisible := False;
        GRem := THGrid(GetControl('GREM'));
        If Grem <> Nil then
        begin
                GRem.ColAligns[2] := taRightJustify;
                GRem.OnCellEnter := GrilleCellEnter;
                GRem.ColFormats[2] := '# ##0';
                GRem.OnCellExit := GrilleCellExit;
                GRem.OnElipsisClick := GrilleElipsisClick;
                GRem.HideSelectedWhenInactive:=true;
        end;
        SetControlProperty('PE3_UGMSA','MaxLength',2);  //PT1

  //DEB PT4
  for i := 1 to 4 do
    THValComboBox(GetControl('TYPEREM' + IntToStr(i))).OnEnter := cbComboChange;  //PT9 à la place de l'évènement
                                                                                  //OnChange car sinon indique qu'il
                                                                                  //y a eu des modifs au chargement
  for i := 1 to 4 do
  begin
    THEdit(GetControl('TAUXMAJO' + IntToStr(i))).OnEnter := cbValChange;     //PT9
    THEdit(GetControl('NBHEURES' + IntToStr(i))).OnEnter := cbValChange;     //PT9
    THEdit(GetControl('MONTANTREM' + IntToStr(i))).OnEnter := cbValChange;   //PT9
  end;
  //FIN PT4
end ;

procedure TOM_MSAPERIODESPE31.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_MSAPERIODESPE31.OnCancelRecord ;
begin
  Inherited ;
        remplirGrilleRemuneration;
end ;

procedure TOM_MSAPERIODESPE31.RemplirGrilleRemuneration ;
Var i : Integer;
    TAutrePeriode : Tob;
    Q : TQuery;
begin
        For i := 1 to GRem.RowCount -1 do
        begin
                GRem.CellValues[0,i] := '';
                GRem.CellValues[1,i] := '';
                GRem.CellValues[2,i] := '0';
        end;
        For i := 1 to 5 do
        begin
                GRem.CellValues[0,i] := GetField('PE3_ELTCALCMSA'+ IntToStr(i));
                If GetField('PE3_MONTANTELT'+ IntToStr(i)) <> 0 then GRem.CellValues[2,i] := IntToStr(GetField('PE3_MONTANTELT'+ IntToStr(i)));
                If GetField('PE3_ELTCALCMSA'+ IntToStr(i)) <> '' then GRem.CellValues[1,i] := RechDom('PGMSATYPEELEMENTCALC',GetField('PE3_ELTCALCMSA'+ IntToStr(i)),False);
        end;
        Q := OpenSQL('SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="'+GetField('PE3_SALARIE')+'" AND '+
        'PE3_DATEDEBUT="'+UsDateTime(GetField('PE3_DATEDEBUT'))+'" AND PE3_DATEFIN="'+UsDateTime(GetField('PE3_DATEFIN'))+'" '+
        'AND PE3_ORDRE=1',True);
        TAutrePeriode := Tob.Create('',Nil,-1);
        TAutrePeriode.LoadDetailDB('table','','',Q,False);
        Ferme(Q);
        If TAutrePeriode.Detail.Count >0 then
        begin
                For i := 6 to 10 do
                begin
                        GRem.CellValues[0,i] := TAutrePeriode.Detail[0].GetValue('PE3_ELTCALCMSA'+ IntToStr(i-5));
                        GRem.CellValues[2,i] := IntToStr(TAutrePeriode.Detail[0].GetValue('PE3_MONTANTELT'+ IntToStr(i-5)));
                        If TAutrePeriode.Detail[0].GetValue('PE3_ELTCALCMSA'+ IntToStr(i-5)) <> '' then GRem.CellValues[1,i] := RechDom('PGMSATYPEELEMENTCALC',TAutrePeriode.Detail[0].GetValue('PE3_ELTCALCMSA'+ IntToStr(i-5)),False);
                end;
        end;
        TAutrePeriode.Free;
        For i := 1 to GRem.RowCount -1 do
        begin
                If GRem.CellValues[0,i] = '' then GRem.CellValues[2,i] := '';
        end;
      //  If GRem.CellValues[0,5] <> '' then GRem.ScrollBars := ssVertical
      //  else GRem.ScrollBars := ssNone;

      //DEB PT4
      for i := 0 to 3 do
      begin
        Q := OpenSQL('SELECT PE3_MONTANTELT1,PE3_MONTANTELT2,PE3_MONTANTELT3,PE3_MONTANTELT4 FROM MSAPERIODESPE31 ' +
          ' WHERE PE3_SALARIE="'+GetField('PE3_SALARIE')+'" AND ' +
          'PE3_DATEDEBUT="'+UsDateTime(GetField('PE3_DATEDEBUT'))+'" AND PE3_DATEFIN="'+UsDateTime(GetField('PE3_DATEFIN'))+'" '+
          'AND PE3_ORDRE=' + IntToStr(i + 100),True);
        if not Q.Eof then
        begin
          SetControlText('TYPEREM' + IntToStr(i + 1),FloatToStr(Q.FindField('PE3_MONTANTELT1').AsFloat));
          SetControlText('TAUXMAJO' + IntToStr(i + 1),FloatToStr(Q.FindField('PE3_MONTANTELT2').AsFloat));
          SetControlText('NBHEURES' + IntToStr(i + 1),FloatToStr(Q.FindField('PE3_MONTANTELT3').AsFloat / 100));
          SetControlText('MONTANTREM' + IntToStr(i + 1),FloatToStr(Q.FindField('PE3_MONTANTELT4').AsFloat));
        end;
        Ferme(Q);
      end;
      //FIN PT4
end;

procedure TOM_MSAPERIODESPE31.GrilleElipsisClick(Sender: TObject);
var SWhere : String;
begin
        sWhere:= 'CO_TYPE="PMZ"';
        LookUpList (GRem,'Eléments de calcul','COMMUN','CO_CODE','CO_LIBRE',swhere,'CO_CODE',TRUE,-1);

end;

procedure TOM_MSAPERIODESPE31.GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
        If GRem.Col = 0 then GRem.ElipsisButton := True
        else GRem.ElipsisButton := False;
        SetField('PE3_SALARIE',GetField('PE3_SALARIE'));
        ForceUpdate;
end;

procedure TOM_MSAPERIODESPE31.GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
        If ACol = 0 then GRem.CellValues[1,GRem.Row] := RechDom('PGMSATYPEELEMENTCALC',GRem.CellValues[0,GRem.Row],False);
    //    If (ARow > 4) and (GRem.CellValues[0,GRem.Row] <> '') then GRem.ScrollBars := ssVertical
    //    else GRem.ScrollBars := ssNone;
end;

procedure TOM_MSAPERIODESPE31.ExitEdit(Sender: TObject);  //PT3
var edit : thedit;
    Salarie:String;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<10) and (isnumeric(edit.text)) then
       begin
       Salarie:=AffectDefautCode(edit,10);
       edit.text:=Salarie;
       SetField('PE3_SALARIE',Salarie);
       end;
end;

//DEB PT4
procedure TOM_MSAPERIODESPE31.cbComboChange(Sender: TObject);
begin
       SetField('PE3_SALARIE',GetField('PE3_SALARIE'));
       ForceUpdate ;   //PT9 en CWAS, le SetField ne suffit pas pour dire qu'il y a eu modification
end;

procedure TOM_MSAPERIODESPE31.cbValChange(Sender: TObject);
begin
       SetField('PE3_SALARIE',GetField('PE3_SALARIE'));
       ForceUpdate ;   //PT9 en CWAS, le SetField ne suffit pas pour dire qu'il y a eu modification
end;
//FIN PT4

//DEB PT7
procedure TOM_MSAPERIODESPE31.OnAfterDeleteRecord;
begin
  inherited;
  ExecuteSQL('DELETE FROM MSAPERIODESPE31' +
    ' WHERE PE3_SALARIE="' + MemSalarie + '"' +
    ' AND PE3_DATEDEBUT="' + USDATETIME(MemDateDebut) + '"' +
    ' AND PE3_DATEFIN="' + USDATETIME(MemDateFin) + '"');
end;
//FIN PT7

Initialization
  registerclasses ( [ TOM_MSAPERIODESPE31 ] ) ;
end.

