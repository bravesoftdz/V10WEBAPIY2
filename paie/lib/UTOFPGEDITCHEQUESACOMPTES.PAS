{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGEDITCHEQUESACOMPTES ()
Mots clefs ... : TOF;PGEDITCHEQUESACOMPTES
*****************************************************************
PT1 : 10/12/2001 : JL 563: Changement du message en cas d'édition effectué, une question est désormais posée
                           et le update de PSD_TOPREGLE est effectué lors de l'impression de l'état.
PT2 : 13/08/2004 : JL V_50 : FQ 11498 Ne pas éditer les montants négatifs
PT3 : 20/04/2005 : JL V_60 : FQ 12204 Ajout clause order by
PT4 : 25/01/2007 : FC V_80 : Mise en place filtrage des habilitations/poupulations
PT5 : 24/09/2009 JS Edition des chèques salaires impossible dans la gestion des acomptes
}
Unit UTOFPGEDITCHEQUESACOMPTES ;

Interface

Uses

 StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,UTOF,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtEtat,QRS1,
{$ELSE}
     UTOB,UtileAgl,eQrs1,
{$ENDIF}
     Hctrls,PgEditOutils,PgEditOutils2,HMsgBox,HTB97,HPdfviewer,ParamDat,EntPaie,HEnt1,
     PGoutils,PGoutils2;


Type
  TOF_PGEDITCHEQUESACOMPTES = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure Change(Sender: TObject); //PT5
    private
    procedure DateElipsisclick(Sender: TObject);
    procedure VerifDate(Sender:TObject);
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation

uses P5Def; //PT4

procedure TOF_PGEDITCHEQUESACOMPTES.OnUpdate ;
var SQL,SDate,Etab1,Etab2,Sal1,Sal2,StOrder: string;
    Date:TDateTime;
    reponse : Word;
begin
  Inherited ;
If GetCheckBoxState('CTRIALPHA') = CbChecked then StOrder := ' ORDER BY PSA_LIBELLE,PSD_SALARIE'  //PT3
else StOrder := ' ORDER BY PSD_SALARIE';
SDate:=GetControlText('DATE');
If Not IsValidDate(SDate) Then
   begin
   LastError:=1;
   LasterrorMsg:='La date saisie n"est pas valide';
   exit;
    end;
Date:=StrToDateTime(SDate);
Etab1:=GetControlText('PSD_ETABLISSEMENT');
Etab2:=GetControlText('PSD_ETABLISSEMENT_');
Sal1:=GetControlText('PSD_SALARIE');
Sal2:=GetControlText('PSD_SALARIE_');
   If ExisteSQL('SELECT PSD_SALARIE FROM HISTOSAISRUB LEFT JOIN SALARIES ON PSD_SALARIE=PSA_SALARIE'+
   ' WHERE  PSD_TOPREGLE="X" AND  PSD_DATEDEBUT="'+UsDateTime(Date)+'" AND PSD_ETABLISSEMENT>="'+Etab1+'" AND PSD_ETABLISSEMENT<="'+Etab2+'"'+
      ' AND PSD_SALARIE>="'+Sal1+'" AND PSD_SALARIE<="'+Sal2+'" AND PSA_PAIACOMPTE="002" AND PSD_ORIGINEMVT="ACP"') Then
      begin       //PT1
      reponse:=HShowMessage('1;Edition des chèques acomptes du '+SDate+';Cette édition a déja été efféctuée, #13#10#13#10'+
      'Voulez-vous éditer de nouveau tous les chèques pour cette date ? #13#10 (si vous cliquez sur Non, seuls les paiements non effectués pour ce jour seront édités.);Q;YN;Y;N;;;','','');
      If reponse = 6 then
         begin
         SQL:='SELECT PSD_SALARIE,PSD_DATEDEBUT,PSD_MONTANT,PSA_PRENOM,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_CIVILITE FROM HISTOSAISRUB LEFT JOIN SALARIES ON PSD_SALARIE=PSA_SALARIE'+
         ' WHERE PSD_MONTANT>0 AND PSD_ETABLISSEMENT>="'+Etab1+'" AND PSD_ETABLISSEMENT<="'+Etab2+'" AND PSD_DATEDEBUT="'+UsDateTime(Date)+'" AND PSA_PAIACOMPTE="002"'+ //pt2
         ' AND PSD_SALARIE>="'+Sal1+'" AND PSD_SALARIE<="'+Sal2+'" AND PSD_ORIGINEMVT="ACP"';
         //DEB PT4
           if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
             SQL := SQL + ' AND ' + MonHabilitation.LeSQL;
         //FIN PT4
         TFQRS1 (Ecran).WhereSQL:=SQL+StOrder; //PT4
         end;
      If reponse=7 then
         begin
         SQL:='SELECT PSD_SALARIE,PSD_DATEDEBUT,PSD_MONTANT,PSA_PRENOM,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_CIVILITE FROM HISTOSAISRUB LEFT JOIN SALARIES ON PSD_SALARIE=PSA_SALARIE'+
         ' WHERE PSD_MONTANT>0 AND PSD_ETABLISSEMENT>="'+Etab1+'" AND PSD_ETABLISSEMENT<="'+Etab2+'" AND PSD_DATEDEBUT="'+UsDateTime(Date)+'" AND PSA_PAIACOMPTE="002"'+ //PT2
         ' AND PSD_SALARIE>="'+Sal1+'" AND PSD_SALARIE<="'+Sal2+'" AND PSD_ORIGINEMVT="ACP" AND PSD_TOPREGLE<>"X"';
         //DEB PT4
           if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
             SQL := SQL + ' AND ' + MonHabilitation.LeSQL;
         //FIN PT4
         TFQRS1 (Ecran).WhereSQL:=SQL+StOrder; //PT4
         end;
      end
   Else
     begin
     SQL:='SELECT PSD_SALARIE,PSD_DATEDEBUT,PSD_MONTANT,PSA_PRENOM,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_CIVILITE FROM HISTOSAISRUB LEFT JOIN SALARIES ON PSD_SALARIE=PSA_SALARIE'+
     ' WHERE PSD_MONTANT>0 AND PSD_ETABLISSEMENT>="'+Etab1+'" AND PSD_ETABLISSEMENT<="'+Etab2+'" AND PSD_DATEDEBUT="'+UsDateTime(Date)+'" AND PSA_PAIACOMPTE="002"'+ //PT2
     ' AND PSD_SALARIE>="'+Sal1+'" AND PSD_SALARIE<="'+Sal2+'" AND PSD_ORIGINEMVT="ACP"';
     //DEB PT4
       if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
         SQL := SQL + ' AND ' + MonHabilitation.LeSQL;
     //FIN PT4
     TFQRS1 (Ecran).WhereSQL:=SQL+StOrder; //PT4
     end;
// La requête de MAJ, est maintenant dans l'état.
{    Q:=OpenSQL('SELECT PSD_SALARIE,PSD_DATEDEBUT FROM HISTOSAISRUB LEFT JOIN SALARIES ON PSD_SALARIE=PSA_SALARIE'+
               ' WHERE PSA_PAIACOMPTE="002" AND PSD_DATEDEBUT="'+UsDateTime(Date)+'"'+
               ' AND PSD_ETABLISSEMENT>="'+Etab1+'" AND PSD_ETABLISSEMENT<="'+Etab2+'"'+
               ' AND PSD_SALARIE>="'+Sal1+'" AND PSD_SALARIE<="'+Sal2+'" AND PSD_ORIGINEMVT="ACP"',True);
   While Not Q.Eof Do
         begin
         Salarie:=Q.FindField('PSD_SALARIE').AsString;
         DateDeb:=Q.FindField('PSD_DATEDEBUT').AsDateTime;
         ExecuteSQL('UPDATE HISTOSAISRUB SET PSD_TOPREGLE="X" WHERE PSD_SALARIE="'+Salarie+'" AND PSD_DATEDEBUT="'+UsDateTime(DateDeb)+'" AND PSD_ORIGINEMVT="ACP"') ;
         Q.Next;
         end;
   Ferme(Q); }
end ;


procedure TOF_PGEDITCHEQUESACOMPTES.OnArgument (S : String ) ;
var Min,Max:String;
    Date,Defaut:THEdit;
    Check: TCheckBox;
begin
  Inherited ;
RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
Defaut:=ThEdit(getcontrol('PSD_SALARIE'));
If Defaut<>nil then Begin Defaut.text:=Min; Defaut.OnExit:=ExitEdit; End;
Defaut:=ThEdit(getcontrol('PSD_SALARIE_'));
If Defaut<>nil then Begin Defaut.text:=Max; Defaut.OnExit:=ExitEdit; End;
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
Defaut:=ThEdit(getcontrol('PSD_ETABLISSEMENT'));
If Defaut<>nil then Defaut.text:=Min;
Defaut:=ThEdit(getcontrol('PSD_ETABLISSEMENT_'));
If Defaut<>nil then Defaut.text:=Max;
  Date:=THEdit(GetControl('DATE'));
  If Date<>NIL Then
    begin
    Date.OnElipsisClick := DateElipsisclick;
    Date.OnExit := VerifDate;
    end;
//debut PT5
Check := TCheckBox(GetControl('Fetat'));
  if Check <> nil then
   Check.OnClick := Change;
//fin PT5
end ;

procedure TOF_PGEDITCHEQUESACOMPTES.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCHEQUESACOMPTES.VerifDate(Sender:TObject);
var THDate:THEdit;
begin
THDate:=THEdit(GetControl('DATE'));
If Not IsValidDate(THDate.text) Then
   begin
   PGIBox(''''+THDate.text+''' n''est pas une date correcte','CEGID PAIE S5');
   THDate.SetFocus ;
   exit;
   end ;
end;

procedure TOF_PGEDITCHEQUESACOMPTES.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

//PT5
procedure TOF_PGEDITCHEQUESACOMPTES.Change(Sender: TObject);
begin
If GETCONTROLTEXT('Fetat') ='PCS' then
begin
 PGIBox('Edition des chèques salaires impossible dans la gestion des acomptes', ecran.caption);
 Setcontroltext('Fetat', 'PCA')
end
end;
//PT5
Initialization
  registerclasses ( [ TOF_PGEDITCHEQUESACOMPTES ] ) ; 
end.

