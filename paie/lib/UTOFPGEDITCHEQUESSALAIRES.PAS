{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGEDITCHEQUESSALAIRES ()
Mots clefs ... : TOF;PGEDITCHEQUESSALAIRES
*****************************************************************
PT1 : 10/12/2001 : JL 563: Changement du message en cas d'édition effectué, une question est désormais posée
                           et le update de PPU_TOPREGLE est effectué lors de l'impression de l'état.
PT2 : 10/09/2003 : JL V_42 : FQ 10736 modif pour gestion onglet avancé
PT3 : 13/08/2004 : JL V_50 : FQ 11498 Ne pas éditer les montants négatifs
PT4 : 20/04/2005 : JL V_60 : FQ 12204 Ajout clause order by
PT5 : 24/09/2009 JS Edition des chèques acomptes impossible dans le règlements par chèques
}
Unit UTOFPGEDITCHEQUESSALAIRES ;

Interface

uses StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,UTOF,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtEtat,QRS1,
{$ELSE}
     UTOB,UtileAgl,eQRS1,
{$ENDIF}
     Hctrls,PgEditOutils,PgEditOutils2,HMsgBox,HTB97,HPdfviewer,ParamDat,EntPaie,HEnt1,HQry,
     PGoutils,PGoutils2;

Type
  TOF_PGEDITCHEQUESSALAIRES = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure Change(Sender: TObject); //PT5
    private
    procedure DateElipsisclick(Sender: TObject);
    procedure VerifDate(Sender:TObject);
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation


procedure TOF_PGEDITCHEQUESSALAIRES.OnUpdate ;
var SQL,SDate,StWhere,StOrder : string;
    Date:TDateTime;
    reponse : Word;
    Pages : TPageControl;
begin
  Inherited ;
If GetCheckBoxState('CTRIALPHA') = CbChecked then StOrder := ' ORDER BY PPU_LIBELLE,PPU_SALARIE'  //PT4
else StOrder := ' ORDER BY PPU_SALARIE';
Pages := TPageControl(GetControl('Pages'));
StWhere := RecupWhereCritere(Pages);                                  //PT2
SDate:=GetControlText('DATE');
If Not IsValidDate(SDate) Then
   begin
   LastError:=1;
   LasterrorMsg:='La date saisie n"est pas valide';
   exit;
   end;
Date:=StrToDateTime(SDate);
If StWhere <> '' then StWhere := StWhere + ' AND PPU_PGMODEREGLE="002"' //PT2
else StWhere := 'WHERE PPU_PGMODEREGLE="002"';
If ExisteSQL('SELECT PPU_SALARIE FROM PAIEENCOURS '+StWhere+' AND PPU_TOPREGLE="X"') Then
      begin                //PT1
      reponse:=HShowMessage('1;Edition des chèques salaires du '+SDate+';Cette édition a déja été efféctuée, #13#10#13#10'+
                                  'Voulez-vous éditer de nouveau tous les chèques pour cette date ? #13#10 (si vous cliquez sur Non, seuls les paiements non effectués pour ce jour seront édités.);Q;YN;Y;N;;;','','');
      if reponse = 6 then
         begin
         SQL:='SELECT PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CNETAPAYER,PPU_DATEDEBUT,PPU_DATEFIN,PSA_CIVILITE FROM PAIEENCOURS LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE '+StWhere+' AND PPU_PAYELE="'+UsDateTime(Date)+'" '+
         'AND PPU_CNETAPAYER>0'; //PT3
         TFQRS1 (Ecran).WhereSQL:=SQL+StOrder; //PT4
         end;
      if reponse= 7 then
         begin
         SQL:='SELECT PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CNETAPAYER,PPU_DATEDEBUT,PPU_DATEFIN,PSA_CIVILITE FROM PAIEENCOURS LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE '+StWhere+' AND PPU_TOPREGLE<>"X" AND PPU_PAYELE="'+UsDateTime(Date)+'" '+
         'AND PPU_CNETAPAYER>0'; //PT3
         TFQRS1 (Ecran).WhereSQL:=SQL+StOrder; //PT4
         end;
      end
      Else
          begin
         SQL:='SELECT PPU_SALARIE,PPU_LIBELLE,PPU_PRENOM,PPU_CNETAPAYER,PPU_DATEDEBUT,PPU_DATEFIN,PSA_CIVILITE FROM PAIEENCOURS LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE '+StWhere+' AND PPU_PAYELE="'+UsDateTime(Date)+'" '+
         'AND PPU_CNETAPAYER>0'; //PT3
         TFQRS1 (Ecran).WhereSQL:=SQL+StOrder; //PT4
          end;
// La requête de MAJ, est maintenant dans l'état.
{Q:=OpenSQL('SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_PGMODEREGLE="002" AND PPU_PAYELE="'+UsDateTime(Date)+'"'+
' AND PPU_ETABLISSEMENT>="'+Etab1+'" AND PPU_ETABLISSEMENT<="'+Etab2+'"'+
' AND PPU_SALARIE>="'+Sal1+'" AND PPU_SALARIE<="'+Sal2+'" and PPU_TOPREGLE<>"X"',True);
While Not Q.Eof Do
      begin
      Salarie:=Q.FindField('PPU_SALARIE').AsString;
      DateDeb:=Q.FindField('PPU_DATEDEBUT').AsDateTime;
      DateFin:=Q.FindField('PPU_DATEFIN').AsDateTime;
      ExecuteSQL('UPDATE PAIEENCOURS SET PPU_TOPREGLE="X" WHERE PPU_SALARIE="'+Salarie+'" AND PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'" AND PPU_DATEFIN="'+UsDateTime(DateFin)+'" AND PPU_PGMODEREGLE="002"') ;
      Q.Next;
      end;
Ferme(Q);   }
end ;


procedure TOF_PGEDITCHEQUESSALAIRES.OnArgument (S : String ) ;
var  Min,Max:String;
    Date,Defaut:THEdit;
    Check: TCheckBox;
begin
  Inherited ;
RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
Defaut:=ThEdit(getcontrol('PPU_SALARIE'));
If Defaut<>nil then Begin Defaut.text:=Min; Defaut.OnExit:=ExitEdit; End;
Defaut:=ThEdit(getcontrol('PPU_SALARIE_'));
If Defaut<>nil then Begin Defaut.text:=Max; Defaut.OnExit:=ExitEdit; End;
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
Defaut:=ThEdit(getcontrol('PPU_ETABLISSEMENT'));
If Defaut<>nil then Defaut.text:=Min;
Defaut:=ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
If Defaut<>nil then Defaut.text:=Max;
  Date:=THEdit(GetControl('DATE'));
  If Date<>NIL Then
    begin
    Date.OnElipsisClick := DateElipsisclick;
    Date.OnExit := VerifDate;
    end;
//PT5
Check := TCheckBox(GetControl('Fetat'));
  if Check <> nil then
   Check.OnClick := Change;
//PT5
end ;


procedure TOF_PGEDITCHEQUESSALAIRES.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCHEQUESSALAIRES.VerifDate(Sender:TObject);
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

procedure TOF_PGEDITCHEQUESSALAIRES.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

//debut PT5
procedure TOF_PGEDITCHEQUESSALAIRES.Change(Sender: TObject);
begin
If GETCONTROLTEXT('Fetat') ='PCA' then
begin
 PGIBox('Edition des chèques acomptes impossible dans le règlements par chèques', ecran.caption);
 Setcontroltext('Fetat', 'PCS')
end
end;
//fin PT5
Initialization
  registerclasses ( [ TOF_PGEDITCHEQUESSALAIRES ] ) ; 
end.
