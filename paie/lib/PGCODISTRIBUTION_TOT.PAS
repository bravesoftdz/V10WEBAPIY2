{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 29/04/2003
Modifié le ... : 06/05/2003
Description .. : Source TOT de la TABLETTE : PGCODISTRIBUTION
Suite ........ : (PGCODISTRIBUTION)
Suite ........ : le code distribution sert à constituer les paquets composant
Suite ........ : le colis livré. C'est l'entreprise qui définit ces codes
Mots clefs ... : TOT;PGCODISTRIBUTION
*****************************************************************}
{
 PT1  MF  23/02/2005  V_60 : Pour le fournisseur NATEXIS le code distribution
                             doit être numérique
 PT2  MF  25/02/2005  V_60 : 1-Avant de suprimer un code distribution, il faut
                             vérifier son utilisation dans DEPORTSAL (salariés
                             affectés à ce code) et désaffecter ces salariés
                             avant de supprimer le code distribution
                             2-Pour SODEXHO le code distribution (Abrégé) est
                             limité à 10 caractères maximum.
                             Pour NATEXIS le code distribution est de 7 caractères.
 PT3 MF   07/07/2006  V_70 : FQ 13348 :Pour NATEXIS le code distribution est
                             limité à 7 caractères maximum.
 PT4 MF   22/01/2007  V_720  Nouveau fournisseur : ACCOR. On renseigne le CC_LIBRE
                             Il alimente le Niveau interne 1 
}
Unit PGCODISTRIBUTION_TOT ;

Interface

Uses
//unused     StdCtrls, 
     Controls,
     Classes,
     PgOutils,
     P5Util,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ELSE}
     Utob,
     MainEAGL,
{$ENDIF}
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97, // PT4
     PgOutils2, // PT4
     ParamSoc,
     UTOT ; 

Type
  TOT_PGCODISTRIBUTION = Class ( TOT )
// d PT1
    private
    Fournisseur                        : string;
// f PT1
    public
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnComplement (Prefixe, Tipe, Code: string);        override ;
    procedure OnArgument (S : String ) ; override ; 
  end ;

Implementation

procedure TOT_PGCODISTRIBUTION.OnNewRecord () ;
begin
  Inherited ;
end ;

procedure TOT_PGCODISTRIBUTION.OnDeleteRecord () ;
// d PT2
var
  st                                  : string;
  Q                                   : TQuery;
  nbre                                : integer;
  SQL                                 :String;
// f PT2
begin
  Inherited ;
// d PT2-1
  st := 'SELECT Count (*) NBRE FROM DEPORTSAL '+
        'WHERE PSE_DISTRIBUTION = "' + GetField('CC_ABREGE')+'"';
  Q := OpenSql (st, TRUE) ;
  if not Q.EOF then
    nbre := Q.FindField ('NBRE').AsInteger
  else
    nbre := 0;
  Ferme (Q);

  if nbre <> 0 then
  {suppression impossible}
  begin
    PgiBox('Attention! Ce code distribution a été affecté à '+
                    IntToStr(nbre)+' '+' salariés'+
                    '#13#10Vous ne pouvez le supprimer!','');
  end ;
  if (nbre <> 0) then
  begin
    if (PgiAsk('Voulez- vous procéder à la désaffectation de ces salariés','') <> mrYes) then
    begin
      LastError := 1;
      LastErrorMsg:='Ce code de distribution ne peut pas être supprimé.';
    end
    else
    begin
      SQL:='UPDATE DEPORTSAL SET PSE_TICKETREST="-",PSE_DISTRIBUTION="" '+
           'WHERE PSE_DISTRIBUTION="'+GetField('CC_ABREGE')+'"';
    Case PGIAskCancel('Confirmez-vous la la désaffectation des salariés ?','') of
      mrYes :
      begin
        try
        begintrans;
          ExecuteSQL(SQL);
        Committrans;
        except
          Rollback;
          PGIBox('! Erreur maj table DEPORTSAL','');
        end;
      end;
      mrNo :
      begin
       LastError := 1;
       exit;
      end;
      mrCancel :
      begin
        LastError := 1;
        exit;
      end;
     end;
    end;
  end;
// f PT2
end ;

procedure TOT_PGCODISTRIBUTION.OnUpdateRecord () ;
var
  ExisteCod : Boolean;
  NomChamp: array[1..2] of Hstring;
  ValChamp: array[1..2] of variant;

begin
  Inherited ;
  NomChamp[1]:='CC_TYPE'; ValChamp[1]:='PCD';;
  NomChamp[2]:='CC_ABREGE';  ValChamp[2]:=GetField ('CC_ABREGE');
  ExisteCod:=RechEnrAssocier('CHOIXCOD',NomChamp,ValChamp);
  if ExisteCod=TRUE then
  begin
    LastError:=1;
    LastErrorMsg:='Attention! Ce code de distribution existe déjà,'+
                 '#13#10 changer l''abrégé!';
  end;
// d PT1
  if (Fournisseur = '002') and (not Isnumerique(GetField ('CC_ABREGE'))) then
  begin
  // NATEXIS : code distribution numérique
    LastError:=1;
    LastErrorMsg:='Attention! L''abrégé doit être numérique,'+
                 '#13#10 changer l''abrégé!';
  end;
// f PT1
// d PT2-2
  if (Fournisseur ='001') and (length(GetField('CC_ABREGE')) > 10) then
  begin
  // SODEXHO : code distribution sur 10 caractères maxi
    LastError:=1;
    LastErrorMsg:='Attention! L''abrégé ne doit pas être de plus de 10 caratères,'+
                 '#13#10 changer l''abrégé!';
  end;
//PT3  if (Fournisseur ='002') and (length(GetField('CC_ABREGE')) <> 7) then
  if (Fournisseur ='002') and (length(GetField('CC_ABREGE')) > 7) then
  begin
  // NATEXIS : code distribution sur 7 caractères maxi
    LastError:=1;
    LastErrorMsg:='Attention! L''abrégé ne doit pas être de plus de 7 caratères,'+
                 '#13#10 changer l''abrégé!';
  end;

// f PT2-2

end ;

procedure TOT_PGCODISTRIBUTION.OnAfterUpdateRecord () ;
begin
  Inherited ;
end ;

procedure TOT_PGCODISTRIBUTION.OnClose () ;
begin
  Inherited ;
end ;

procedure TOT_PGCODISTRIBUTION.OnArgument (S : String ) ;
begin
  Inherited ;
  Fournisseur := GetParamSocSecur('SO_PGTYPECDETICKET','');     //PT1
// d PT4
  if (Fournisseur = '003') then
  // ACCOR
    OKComplement := True
  else
  // SODEXHO, NATEXIS
    OKComplement := False;
// f PT4
end ;

// d PT4
// Lancement de la fiche de saisie du CC_LIBRE
procedure TOT_PGCODISTRIBUTION.OnComplement(Prefixe, Tipe, Code: string);
var
   QQ : TQuery;
   CClibre : string;
begin
  inherited;

  QQ := opensql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PCD" AND CC_CODE="'+Code+'"',TRUE);
  if not QQ.EOF then
  begin
    CClibre :=  QQ.findfield('CC_LIBRE').asstring;
  end;
  ferme (QQ);

  if (Code = '') or (GetField('CC_LIBELLE') = '') then Exit;

  AglLanceFiche('PAY', 'LIBRE', '', '', 'CODISTRIBUTION;'+Code+';'+CClibre);

end;
// f PT4

Initialization
  registerclasses ( [ TOT_PGCODISTRIBUTION ] ) ;
end.
