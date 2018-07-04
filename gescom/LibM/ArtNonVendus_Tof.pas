{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 15/02/2002
Modifié le ... : 17/09/2002
Description .. : Source TOF de la FICHE : ARTNONVENDUS ()
Suite ........ : Edtion des articles non vendus =
Suite ........ :  Stock physique > 0 et Ventes FFO = 0
Mots clefs ... : TOF;ARTNONVENDUS
*****************************************************************}
Unit ArtNonVendus_Tof ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,
//     eQRS1,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,QRS1,HQry,UtilPGI ; 

Type
  TOF_ArtNonVendus = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_ArtNonVendus.OnLoad ;
var
  stWhere1, stWhere2 : string ;
  mSql : TQuery;
begin
  Inherited ;
  //
  // Récupération des critères
  //
  stWhere1 := MultiComboInSQL(GetControlText('ETABLISSEMENT'));
  if stWhere1 <> '' then
  begin
    if stWhere1 <> TraduireMemoire( '"<<Tous>>"' ) then
    begin
      mSql := OpenSQL( 'SELECT ET_DEPOTLIE FROM ETABLISS WHERE ET_ETABLISSEMENT IN ('+stWhere1+')', True);
      stWhere1 := '';
      while not mSql.Eof do
      begin
        if stWhere1 = '' then
          stwhere1 := stWhere1 + mSql.FindField('ET_DEPOTLIE').AsString
        else
          stwhere1 := stWhere1 + ';'+ mSql.FindField('ET_DEPOTLIE').AsString;
        mSql.Next;
      end;
      stWhere1 := 'GQ_DEPOT IN ("' + stWhere1 + '") AND';
      stWhere1 := StringReplace (stWhere1, ';', '","', [rfReplaceAll]) ;
      stWhere2 := stWhere1;
      stWhere2 := StringReplace (stWhere2, 'GQ_DEPOT', 'GL_DEPOT', [rfReplaceAll]) ;
      Ferme( mSql );
    end
    else
    begin
      stWhere1 := '';
      stWhere2 := '';
    end;
  end;
  //  Initialisation des variables pour la requête de l'état
  //
  SetControlText('XX_VARIABLE3', stWhere1) ;
  SetControlText('XX_VARIABLE4', stWhere2) ;
  SetControlText('XX_VARIABLE1', USDateTime(StrToDate(GetControlText('GL_DATEPIECE')))) ;
  SetControlText('XX_VARIABLE2', USDateTime(StrToDate(GetControlText('GL_DATEPIECE_')))) ;
end ;

procedure TOF_ArtNonVendus.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  // Affichage du libellé des familles
  //
  THLabel(TForm(Ecran).FindComponent('TGL_FAMILLENIV1')).Caption := RechDom('GCLIBFAMILLE','LF1',False) ;
  THLabel(TForm(Ecran).FindComponent('TGL_FAMILLENIV2')).Caption := RechDom('GCLIBFAMILLE','LF2',False) ;
  THLabel(TForm(Ecran).FindComponent('TGL_FAMILLENIV3')).Caption := RechDom('GCLIBFAMILLE','LF3',False) ;
  THLabel(TForm(Ecran).FindComponent('TGL_COLLECTION')).Caption := RechDom('GCZONELIBRE','AS0',False) ;
end ;

Initialization
  registerclasses ( [ TOF_ArtNonVendus ] ) ;
end.
