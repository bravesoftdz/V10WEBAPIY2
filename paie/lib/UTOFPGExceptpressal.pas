{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXCEPTPRESSAL_MUL ()
                 Saisie des exceptions de présence pour un salarié
Mots clefs ... : UTOFPGEXCEPTPRESSAL
*****************************************************************
PT1  20/07/2007  FLO         Gestion des saisies et suppressions groupées
}
Unit UTOFPGExceptpressal ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     HDB,
     fe_main,
{$else}
     eMul, 
     uTob,
     maineagl,
{$ENDIF}
     forms,
     AglInit ,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     P5def,
     HTB97,
     Hqry,
     UTOF ; 

Type
  TOF_PGExceptpressal = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;

    private
     GblListe , Selectsal: string;
     procedure ActiveWhere (Sender: TObject);
     procedure ChangeListe(Sender: TObject);
     procedure CreationException(Sender : Tobject);
     procedure Exceptionpresence(Sender : Tobject);
     procedure SuppressionException (Sender : TObject);
     procedure GrilleDblClick (Sender : Tobject);
  end ;

Implementation


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /    
Description .. : Chargement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGExceptpressal.OnLoad ;

begin
  Inherited ;
//  ActiveWhere (NIL); //PT1
     ChangeListe(Nil);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /
Description .. : On Argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGExceptpressal.OnArgument (S : String ) ;
var
CBLISTEVAL : THvalcombobox;
num : integer;
Btn : TToolBarButton97;
{$IFNDEF EAGLCLIENT}
Liste:THDBGrid;
{$ELSE}
Liste:THGrid;
{$ENDIF}
begin
  Inherited ;

  // Lnitialise les champs libres
  For Num := 1 to 4 do
     VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
     VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;

  // Gestion des boutons "ouvrir" , "insert"  et "double click"
  {$IFNDEF EAGLCLIENT}
  Liste:=THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste:=THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste <> NIL then Liste.OnDblClick := GrilleDblClick;

  Btn := TToolBarButton97(GetControl('Binsert'));
  if btn <> nil then Btn.Onclick := CreationException;   // Création exception pour un salarié

  Btn := TToolBarButton97(GetControl('BOUVRIR'));
  if btn <> nil then Btn.Onclick := Exceptionpresence;   // Création exception pour une sélection de salariés

  //PT1 - Début
  Btn := TToolBarButton97(GetControl('BDelete'));
  If Btn <> Nil Then Btn.Onclick := SuppressionException;   // Suppression d'une sélection d'exceptions 
  //PT1 - Fin

  // Changement de liste
  SetControlText('CBLISTE','EXC'); //PT1
  Cblisteval := THValComboBox(GetControl('CBLISTE'));
  if Assigned(cblisteval) then Cblisteval.Onchange := ChangeListe;

//  ChangeListe(nil);  //PT1

end ;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... : 23/07/2007 / PT1
Description .. : Active la clause WHERE
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGExceptpressal.ActiveWhere(Sender: TObject);
var
    StWhere   : String;
    datedebut : TDatetime;
begin
     DateDebut := StrToDate(GetControlText('DATEDEBUT'));
     SetControlText('XX_WHERE','');
     StWhere := '';

     StWhere := RecupWhereCritere(GetControl('PAGES') As TPageControl);
     If StWhere <> '' Then
     Begin
          // Suppression du mot clé WHERE
          StWhere := Copy(StWhere, 6, Length(StWhere));
          StWhere := StWhere + ' AND ';
     End;

     StWhere := StWhere + '(PSA_DATESORTIE >="'+UsDateTime(DateDebut)+'" OR PSA_DATESORTIE <="'+UsdateTime(iDate1900)+'" OR' +
                ' PSA_DATESORTIE IS NULL) ';

     If SelectSal <> '' Then Stwhere := Stwhere + ' AND ' + SelectSal;

     If GetControlText('CBLISTE') = 'EXC' Then
          StWhere := StWhere + ' AND PYE_DATEDEBUT >= "'+UsDateTime(DateDebut)+'" ';

     SetControlText('XX_WHERE', StWhere);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... : 20/07/2007 / PT1
Description .. : Changement de liste 
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGExceptpressal.ChangeListe(Sender: TObject);
begin
     // Construit la clause Where
     ActiveWhere(nil);

     // Evite de rafraîchir la liste si c'est la même qui a été resélectionnée
     if GblListe = GetControlText('CBLISTE') then exit;

     // Déselection des lignes
     TFMul(Ecran).FListe.ClearSelected; //PT1

     // Changement de la liste associée
     if GetControlText('CBLISTE')= 'SAL' then
     Begin
          TFMul(Ecran).SetDBListe('PGMULSALARIE');
          //PT1 - Début
          SetControlVisible ('BOuvrir', True);
          SetControlVisible ('BDelete', False);
          //PT1 - Fin
     End
     else if GetControlText('CBLISTE')= 'EXC' then
     Begin
          TFMul(Ecran).SetDBListe('PGEXCEPTPRESSAL');
          //PT1 - Début
          SetControlVisible ('BOuvrir', False);
          SetControlVisible ('BDelete', True);
          //PT1 - Fin
     End;

     // Sauvegarde de la liste en cours
     GblListe := GetControlText('CBLISTE');
     
     // Rafraîchissement de la liste
     If Sender <> Nil Then             //PT1
          TFMul(Ecran).BCherche.Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /
Description .. : Double click
Mots clefs ... :
*****************************************************************}
procedure TOF_PGExceptpressal.GrilleDblClick (Sender : Tobject);
var Q_Mul:THQuery ;
    St : String;
    {$IFDEF EAGLCLIENT}
    Liste:THGrid;
    {$ENDIF}
begin
    {$IFDEF EAGLCLIENT}
    Liste:=THGrid(GetControl('FLISTE'));
    TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
    {$ENDIF}
    Q_Mul:=THQuery(Ecran.FindComponent('Q'));

    // Si liste exception : Modification d'une exception déjà saisie
    if GetControlText('CBLISTE')= 'EXC' then
    begin
     st  := Q_MUL.FindField('PYE_SALARIE').AsString +';' + DateTostr(Q_MUL.FindField('PYE_DATEDEBUT').AsDatetime);
  
     AGLLanceFiche('PAY','EXCEPTPRESSAL','',St,'ACTION=MODIFICATION');
     TFMul(Ecran).BCherche.Click;
    end
    else
    // si liste salariés : saisie d'une exception pour le salarié sélectionné
    begin
      st:= 'ACTION=CREATION' + ';' + Q_MUL.FindField('PSA_SALARIE').AsString + ';' + Q_MUL.Findfield('PSA_LIBELLE').asstring  + ';' +
      Q_MUL.Findfield('PSA_PRENOM').Asstring+';'+StDate1900+';'+StDate1900;
      AGLLanceFiche('PAY','EXCEPTPRESSAL','','',st);
    end;
end;

{***********A.G.L.**********************************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /
Description .. : Insert : Création d'une exception de présence d'un salarié

Mots clefs ... :
****************************************************************************}
procedure TOF_PGExceptpressal.CreationException (Sender : Tobject);

begin
    AGLLanceFiche('PAY','EXCEPTPRESSAL','','','ACTION=CREATION');
    TFMul(Ecran).BCherche.Click;
end;

{***********A.G.L.*****************************************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /    
Description .. : Saisie d'une exception de présence pour les salariés sélectionnés
Mots clefs ... : 
**********************************************************************************}
procedure TOF_PGExceptpressal.Exceptionpresence (sender : Tobject);
var
  Salarie, st : string;
  i           : integer;
begin

     if (TFmul(Ecran).Fliste.NbSelected = 0) and (not TFmul(Ecran).Fliste.AllSelected) then
     begin
          PGIBOX('Aucun salarié sélectionné', Ecran.Caption);
          exit;
     end;

     { Gestion de la sélection des salariés}

     if PgiAsk('Saisie d''une exception de présence pour les salariés sélectionnés. Voulez-vous poursuivre ?', Ecran.caption) = mrYes then
     begin  // D2

          St := '';
          if (TFmul(Ecran).Fliste.AllSelected = TRUE) then
          begin   // D3
               // Si tout est sélectionné
               TFmul(Ecran).Q.First;
               while not TFmul(Ecran).Q.EOF do
               begin
                    Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
                    St := St + ' PSA_SALARIE="' + Salarie + '" OR';
                    TFmul(Ecran).Q.Next;
               end;
               TFMul(Ecran).FListe.ClearSelected;
          end
          else
          // lecture de chaque salarié sélectionné
          begin
               { Composition du clause WHERE pour limiter le mul à ces salariés }
               for i := 0 to TFmul(Ecran).Fliste.NbSelected - 1 do
               begin
                    {$IFDEF EAGLCLIENT}
                         TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
                    {$ENDIF}
                    TFMul(Ecran).Fliste.GotoLeBOOKMARK(i);
                    Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
                    St := St + ' PSA_SALARIE="' + Salarie + '" OR';
               end;
               TFMul(Ecran).FListe.ClearSelected;
          end;

          if St <> '' then St := Copy(St, 1, Length(st) - 2)  ;
          Selectsal :=  St;
          TFMul(Ecran).BCherche.Click;


          { Récupération de la Query pour traitement dans la fiche vierge }
          {$IFDEF EAGLCLIENT}
               if TFMul(Ecran).Fetchlestous then TheMulQ := TOB(Ecran.FindComponent('Q'));
          {$ELSE}
               TheMulQ := THQuery(Ecran.FindComponent('Q'));
          {$ENDIF}

          {Ouverture de la fiche}
          AglLanceFiche('PAY', 'EXCEPTPRESSALGRP', '', '', '');
          
          TheMulQ := nil;
          Selectsal := '';
          TFMul(Ecran).BCherche.Click;
     end; // E2
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 23/07/2007 / PT1
Modifié le ... :   /  /    
Description .. : Suppression groupée d'exceptions
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGExceptpressal.SuppressionException(Sender: TObject);
Var
     Salarie, St, DateDebut : String;
     i, Reponse             : Integer;
begin
     //PT1 - Début
     If (TFmul(Ecran).Fliste.NbSelected = 0) And (Not TFmul(Ecran).Fliste.AllSelected) Then
     begin
          PGIBox(TraduireMemoire('Aucune exception n''est sélectionnée'), Ecran.Caption);
          Exit;
     End;

     If (TFmul(Ecran).Fliste.NbSelected = 1) Then
          Reponse := PgiAsk(TraduireMemoire('Etes-vous sûr de vouloir supprimer l''exception sélectionnée ?'), Ecran.caption)
     Else
          Reponse := PgiAsk(TraduireMemoire('Etes-vous sûr de vouloir supprimer les exceptions sélectionnées ?'), Ecran.caption);

     If Reponse = mrYes Then
     Begin
          St := '';

          // Cas 1 : Tout est sélectionné
          If (TFmul(Ecran).Fliste.AllSelected = True) Then
          Begin
               TFmul(Ecran).Q.First;
               While Not TFmul(Ecran).Q.EOF Do
               Begin
                    Salarie   := TFmul(Ecran).Q.FindField('PYE_SALARIE').AsString;
                    DateDebut := TFmul(Ecran).Q.FindField('PYE_DATEDEBUT').AsString;
                    St := St + ' (PYE_SALARIE="' + Salarie + '" AND PYE_DATEDEBUT="'+UsDateTime(StrToDate(DateDebut))+'") OR';
                    TFmul(Ecran).Q.Next;
               End;
               TFMul(Ecran).FListe.ClearSelected;
          End
          // Cas 2 : Un certain nombre de salariés a été sélectionné
          Else
          Begin
               For i := 0 To TFmul(Ecran).Fliste.NbSelected - 1 Do
               Begin
                    {$IFDEF EAGLCLIENT}
                         TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
                    {$ENDIF}
                    TFMul(Ecran).Fliste.GotoLeBOOKMARK(i);
                    Salarie   := TFmul(Ecran).Q.FindField('PYE_SALARIE').AsString;
                    DateDebut := TFmul(Ecran).Q.FindField('PYE_DATEDEBUT').AsString;
                    St := St + ' (PYE_SALARIE="' + Salarie + '" AND PYE_DATEDEBUT="'+UsDateTime(StrToDate(DateDebut))+'") OR';
               End;
               TFMul(Ecran).FListe.ClearSelected;
          End;

          // Suppression du OR final
          If St <> '' Then St := Copy(St, 1, Length(st) - 2)  ;

          // Exécution de la requête
          ExecuteSQL ('DELETE FROM EXCEPTPRESENCESAL WHERE '+St);

          // Rafraîchissement de la liste
          TFMul(Ecran).BCherche.Click;
      end;
      //PT1 - Fin
end;

Initialization
  registerclasses ( [ TOF_PGExceptpressal ] ) ;
end.
