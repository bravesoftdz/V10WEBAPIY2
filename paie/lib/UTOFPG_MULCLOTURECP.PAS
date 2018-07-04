{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 16/09/2002
Modifié le ... :   /  /
Description .. : Multicritère de cloture des congés payés
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1   15/10/2003 SB V_42 Maj de date modification de la table des salaries pour introduction PAIE
PT2   12/03/2004 SB V_50 FQ 11162 Encodage de la date de cloture erroné si fin fevrier
PT3   22/03/2004 SB V_50 Retrait de l'option de décloture des CP
PT4   27/05/2004 SB V_50 FQ 11308 Double message d'alerte en cloture CP si paies non éffectuées
PT5   27/05/2004 JL V_50 FQ 11392 Message si toutes les paies ne sont pas effectué modifier pour avoir non par défaut
PT6   21/03/2005 SB V_60 FQ 12047 Raffraissement grille en cwas
PT7   18/07/2007 MF V_72 FQ 14558 correction sélection avec bouton "Tout sélectionner"
PT8   12/06/2008 FC V_81 FQ 15536 anomalie clôture en CWAS
}
unit UtofPG_MulClotureCP;

interface
uses Controls,Classes,sysutils,
{$IFDEF EAGLCLIENT}
       eMul,
{$ELSE}
       Mul,
{$ENDIF}
       UTOF,HCtrls,
       HEnt1,HMsgBox,HTB97, 
       Hstatus,ULibEditionPaie;

Type
     TOF_PGMULCLOTURECP = Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure ClotcpClick(Sender : TObject);
     END ;

       Function ClotureCp(Mode, Etab,EtabLibelle,DateCloture,MethodeRel : string) : Boolean;


implementation


{ TOF_PGMULCLOTURECP }

uses P5Def,PGCongesPayes,PGCalendrier;

procedure TOF_PGMULCLOTURECP.OnArgument(Arguments: String);
var ClotCp : TToolBarButton97;
  Where: THEdit; //PT7
  strWhere:String; //PT7
begin
inherited;
  ClotCp := TToolBarButton97 (getcontrol('TBCLOTCP'));
  if Clotcp <> nil then ClotCp.Onclick := ClotCpClick;
  ClotCp := TToolBarButton97 (getcontrol('TBDECLOTCP'));
  if ClotCp <> nil then ClotCp.Onclick := ClotcpClick;
  SetControlVisible('TBDECLOTCP',False); { PT3 }
  //DEB PT7
  if Assigned(MonHabilitation) then
    if (MonHabilitation.Active) and (MonHabilitation.LeSQL <> '') then
    begin
      Where := THEdit(GetControl('XX_WHERE'));
      strWhere := ' ETB_ETABLISSEMENT IN (SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
      if Where <> nil then SetControlText('XX_WHERE', strWhere);
    end;
  //FIN PT7
end;


procedure TOF_PGMULCLOTURECP.ClotcpClick(Sender: TObject);
Var Mode,libelle,Etab,EtabLibelle,DateCloture,MethodeRel,CodeEvt,EtatEvt : string;
i,Tot : integer;
Okok : Boolean;
TEvent : TStringList;
begin
Mode:='';EtatEvt:='OK';  Okok:=True;
If (TFMul(Ecran).FListe=nil) then exit;
if Sender is TToolBarButton97 then
  Begin
  if TToolBarButton97(sender).Name='TBCLOTCP' then
    Begin
    Mode:='CLOTURE';
    Libelle:='Clôture';
    CodeEvt:='012';
    End
  else
    if TToolBarButton97(sender).Name='TBDECLOTCP' then
      Begin
      Mode:='DECLOTURE';
      Libelle:='De-Clôture';
      CodeEvt:='013';
      End;
  End;


if (TFMul(Ecran).FListe.nbSelected=0) and (not TFMul(Ecran).FListe.AllSelected) then
   begin
   MessageAlerte('Aucun élément sélectionné');
   exit;
   end;

if ((TFMul(Ecran).FListe.nbSelected>0) OR (TFMul(Ecran).FListe.AllSelected )) then
   begin
   if TFMul(Ecran).FListe.AllSelected then
     begin
     tot:=TFmul(Ecran).Q.RecordCount;
     TFmul(Ecran).Q.First;
     End
   else Tot:=TFMul(Ecran).FListe.NbSelected;
   InitMove(Tot,'');
   TEvent:=TStringList.Create;
   TEvent.Add(Libelle+' des congés payés des établissements :');
   for i:=0 to Tot-1 do
     BEGIN
     if not TFMul(Ecran).FListe.AllSelected then
     begin
       TFMul(Ecran).FListe.GotoLeBookmark(i);
       {$IFDEF EAGLCLIENT} { PT6 }
// PT7     TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
         TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1); //PT8
       {$ENDIF}
     end;
     MoveCur(False);
     Etab:=TFmul(Ecran).Q.FindField('ETB_ETABLISSEMENT').asstring;
     EtabLibelle:=TFmul(Ecran).Q.FindField('ETB_LIBELLE').asstring;
     DateCloture:=DateToStr(TFmul(Ecran).Q.FindField('ETB_DATECLOTURECPN').AsDateTime);
     MethodeRel:=TFmul(Ecran).Q.FindField('ETB_RELIQUAT').asstring;
     Okok:=ClotureCp(Mode,Etab,EtabLibelle,DateCloture,MethodeRel);
     TEvent.Add(Libelle+' '+Etab+' '+EtabLibelle+' à la date du '+DateCloture+'.');
     if not Okok then begin EtatEvt:='ERR'; Break; end else EtatEvt:='OK';
     if TFMul(Ecran).FListe.AllSelected then TFmul(Ecran).Q.Next;
     END;
   TFMul(Ecran).FListe.ClearSelected;
   TFMul(Ecran).FListe.AllSelected:=False;
   FiniMove;
   if Okok then
      TEvent.Add(Libelle+' des congés payés OK.')
   else
      TEvent.Add('Une erreur est survenue lors de la '+LowerCase(Libelle)+' des congés payés.');
   CreeJnalEvt ('002',CodeEvt,EtatEvt,nil,nil,TEvent);
   if TEvent<>nil then TEvent.free;
   TFMul(Ecran).BCherche.Click;
   end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 16/09/2002
Modifié le ... : 16/09/2002
Description .. : Fonction de clôture des congés payés
Mots clefs ... : PAIE;CP
*****************************************************************}
Function ClotureCp(Mode, Etab,EtabLibelle,DateCloture,MethodeRel : string) : Boolean;
var
FinPeriodeN,FinPeriodeN1, DebPeriodeS,FinPeriodeS : Tdatetime;
Init,aa,mm,jj: word;
begin
result:=False;
if IsValidDate(DateCloture) then
  FinPeriodeN:=StrToDate(DateCloture)
else exit;
IF mode='CLOTURE' then
  BEGIN
  DebPeriodeS := FinPeriodeN+1;
  decodedate(FinPeriodeN,aa,mm,jj);
  FinPeriodeS := PGEncodeDateBissextile(aa+1,mm,jj); { PT2 }
  Init := HShowMessage('2;Congés payés;Attention, vous allez clôturer la période '+
  'de congés payés #13#10et ouvrir la suivante ( du '+datetostr(DebPeriodeS) +
  ' au'+ datetostr(finPeriodeS)+ ') de l''établissement "'+EtabLibelle+'".#13#10 '
  +'Etes vous sûr d''avoir terminé vos paies jusqu''au '+datetostr(FinPeriodeN)+'?#13#10 Voulez-vous continuer?;Q;YN;N;N;','','');  //PT8
  If Init = mrYes then
    begin       { DEB PT4 09/06/2004 }
    if ExisteSql('SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEFIN>="'+USDateTime(DebutDeMois(FinPeriodeN))+'" '+
    'AND PPU_DATEFIN <= "'+USDateTime(FinPeriodeN)+'"') = False then
       if HShowMessage('2;Clôture congés payés;Attention! Vous n''avez pas effectué les paies sur tous les mois #13#10 '+  //PT5
       'de la période congés payés allant du '+DateToStr(DebutDeMois(PlusMois(FinPeriodeN,-11)))+
       ' au '+DateToStr(FinPeriodeN)+'!#13#10 Voulez-vous clôturer?;Q;YN;N;N;','','') = MrNo then  //PT8
          exit; { FIN PT4 }
    if CloturePeriodeCP(FinPeriodeN, etab,MethodeRel) then
      Begin
      ExecuteSql('UPDATE ETABCOMPL SET ETB_DATECLOTURECPN="'+UsDateTime(FinPeriodeS)+'", '+
      'ETB_DATEMODIF="'+UsTime(Now)+'" '+  //PT1
      'WHERE ETB_ETABLISSEMENT="'+Etab+'" ');
      Result:=True;
      End
    Else
      result:=False;
    end;
  END
ELSE
  IF Mode='DECLOTURE' then
    BEGIN
    If DeCloturePeriodeCP(FinPeriodeN,etab,EtabLibelle) then
       Begin
       FinPeriodeN1:=PlusDate(FinPeriodeN,-1,'A');
       ExecuteSql('UPDATE ETABCOMPL SET ETB_DATECLOTURECPN="'+UsDateTime(FinPeriodeN1)+'", '+
      'ETB_DATEMODIF="'+UsTime(Now)+'" '+ //PT1
       'WHERE ETB_ETABLISSEMENT="'+Etab+'" ');
       Result:=True;
       End
    Else
       Result:=False;
    END;
end;

Initialization
registerclasses([TOF_PGMULCLOTURECP]);
end.
 
