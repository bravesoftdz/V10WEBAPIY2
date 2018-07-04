unit P5UtilPaieCompl;

interface

uses Windows,SysUtils,HCtrls,UTob,// Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs;

//  StdCtrls, Hent1,  ComCtrls, HRichEdt, HRichOLE, HMsgBox, HStatus,
//  ExtCtrls, Grids, ImgList, Mask, ClipBrd, richedit, UTOB,  M3VM, ParamSoc, EntPaie;
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables {$ELSE} uDbxDataSet {$ENDIF};//, M3Code, Fe_Main;
{$ELSE}
  MaineAGL;
{$ENDIF}





procedure PgCalculCpBulletin(Action: String; Tob_Rub,Tob_Sal,Tob_Etablissement : Tob;DD,DF : TDateTime) ;


implementation
uses P5util;


procedure PgCalculCpBulletin(Action: String; Tob_Rub,Tob_Sal,Tob_Etablissement : Tob;DD,DF : TDateTime) ;
Var
Q : TQuery;
Salarie,St,StMsg : String;
Tob_PgAbs,T,Ta,Tob_PgPris,Tob_PgSolde,Tob_PgAcquis,Tob_PgPrisPayes : Tob;
i : integer;
AcquisN1, AcquisN, Pris, PrisPayes, ACons, Rel, Restant : Double;
Begin
    FreeAndNil(Tob_PgPris);
    FreeAndNil(Tob_PgSolde);
    FreeAndNil(Tob_PgAcquis);
    FreeAndNil(Tob_PgPrisPayes);

    Salarie := TOB_Sal.GetValeur(iPSA_SALARIE);
{1- Chargement des mouvements en cours }
    St:=   'SELECT * FROM ABSENCESALARIE WHERE PCN_SALARIE = "'+ salarie+'" ' +
                  'AND PCN_TYPEMVT="CPA" AND PCN_TYPERGRPT<>"CLO" AND PCN_CODERGRPT<>-1 '+
(*                  'AND ((PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S" AND ' +
                  '( ( (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS")  ' +
                  ' AND PCN_DATEFIN <> "' + usdatetime(Df) + '")' +
                  ' OR PCN_TYPECONGE="ARR" OR PCN_TYPECONGE="REP" OR ' +
                  ' ( ((PCN_TYPECONGE="AJU") OR (PCN_TYPECONGE="AJP") or(PCN_TYPECONGE="REL")) AND PCN_SENSABS = "+"))) '+
                  'OR ( '+
                  ' ((PCN_TYPECONGE = "PRI" AND PCN_VALIDRESP = "VAL") OR PCN_TYPECONGE = "SLD")' +
                  ' AND PCN_DATEVALIDITE <= "' + usdatetime(Df) + '"' +
                  ' AND PCN_DATEFIN <= "' + usdatetime(10) + '" ) )' +*)
                  ' ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE ';
    Q := OpenSql(St,True);
    if not Q.eof then
      Begin
      Tob_PgAbs := TOB.create('ABSENCE_SALARIE',nil,-1);
      Tob_PgAbs.LoadDetailDb('ABSENCE_SALARIE','','',Q,False);
      Tob_PgPris := TOB.create('ABSENCESALARIE',nil,-1);
      Tob_PgSolde := TOB.create('ABSENCESALARIE',nil,-1);
      Tob_PgAcquis := TOB.create('ABSENCESALARIE',nil,-1);
      Tob_PgPrisPayes := TOB.create('ABSENCESALARIE',nil,-1);
      End;
   Ferme(Q);
  if Assigned(Tob_PgAbs) and (Tob_PgAbs.detail.count>0) then
     Begin
     I := Tob_PgAbs.detail.count-1;
     T:=Tob_PgAbs.Detail[i];
     While Assigned (T) Do
      Begin
      If (T.GetValue('PCN_TYPERGRPT')='PAY') then
         T.ChangeParent(Tob_PgPrisPayes,-1)
      else
        If (T.GetValue('PCN_TYPERGRPT')='PRI') AND (T.GetValue('PCN_DATEVALIDITE') <= Df) then
          T.ChangeParent(Tob_PgPris,-1)
        else
        if (T.GetValue('PCN_TYPERGRPT')='SLD') Then
          T.ChangeParent(Tob_PgSolde,-1)
      else
        if (T.GetValue('PCN_TYPERGRPT')='ACQ') Then
          T.ChangeParent(Tob_PgAcquis,-1);
      Dec(i);
      if i<0 then Break;
      T:=Tob_PgAbs.Detail[i];
      End;
    End;
    if Assigned(Tob_PgPrisPayes) and (Tob_PgPrisPayes.detail.count>0) then
       PrisPayes := Tob_PgPrisPayes.Somme('PCN_JOURS',[''],[''],False);
    if Assigned(Tob_PgAcquis) and (Tob_PgAcquis.detail.count>0) then
      Begin
       AcquisN1 := Tob_PgAcquis.Somme('PCN_JOURS',['PCN_PERIODECP'],['1'],False);
       AcquisN  := Tob_PgAcquis.Somme('PCN_JOURS',['PCN_PERIODECP'],['0'],False);
       Tob_PgAcquis.detail.Sort('PCN_PERIODECP,PCN_TYPECONGE');
      End;

    if Assigned(Tob_PgPris) and (Tob_PgPris.detail.count>0) then
       Begin
       Pris := Tob_PgPris.Somme('PCN_JOURS',[''],[''],False);
       For I := 0 to Tob_PgPris.Detail.count-1 do
         Begin
         T:=Tob_PgPris.Detail[0];
         Pris := T.GetValue('PCN_JOURS');
         if (AcquisN1+AcquisN-PrisPayes) >= Pris then
           Begin
           //Esc-ce que je prends mes conges sur N-1?
           if (AcquisN1-PrisPayes) >= Pris then
              Begin
              Restant := AcquisN1-PrisPayes;
              Ta := Tob_PgAcquis.FindFirst(['PCN_TYPECONGE','PCN_SENSABS'],['REL','+'],False);
              if (Assigned(Ta)) And (Ta.getValue('PCN_JOURS') - PrisPayes > 0) then
                 Begin { Prise cp sur reliquat + }
                 Rel  :=  Ta.getValue('PCN_JOURS');
                 AcquisN1 := AcquisN1  - Rel ;
                 ACons := Rel - PrisPayes ;
                 If Pris <= ACons then
                    Begin
                    StMsg := '#13#10Prise Cp entièrement sur Reliquat + de : '+FloatToStr(Pris);
                    PrisPayes := PrisPayes + Pris;
                    Pris := 0;
                    End
                 else
                    if ACons > 0 then
                      Begin
                      StMsg := '#13#10Prise Cp partiellement sur Reliquat + de : '+FloatToStr(ACons);
                      Pris := Pris - ACons  ;
                      PrisPayes := PrisPayes + ACons ;
                      End
                    else
                      ACons := Rel;
                 End;
              { Prise cp sur les autres acquis }
              if Pris > 0 then
                Begin
                ACons := AcquisN1 - PrisPayes;
                if (AcquisN1>0) and (Pris <= Acons) then { Prise cp sur les acquis de la periode n-1 }
                  Begin
                  StMsg := '#13#10Prise Cp entièrement sur Acquis N1 de : '+FloatToStr(Pris);
                  PrisPayes := PrisPayes + Pris;
                  Pris := 0;
                  End;





                End;

              End;
           End
         else
           Begin
           //PgiBox('Pas assez d''acquis restant','');
           Break;
           End;
         End;


       End;



    FreeAndNil(Tob_PgAbs);
    FreeAndNil(Tob_PgPris);
    FreeAndNil(Tob_PgSolde);
    FreeAndNil(Tob_PgAcquis);
    FreeAndNil(Tob_PgPrisPayes);

{2- Chargement des mouvements antérieurs }


{3- Dispatch des mouvements }


{}



End;


end.
