
#
# INICIALIZACIJA
#
# gsheets security: aplikacijo moraš povezati z google user accountom. Sledi navodilu tukaj https://www.youtube.com/watch?v=TU1znISrAGg
#
# tutorial : https://www.twilio.com/blog/2017/03/google-spreadsheets-ruby.html
# Parametri:
#   preglednica ... v katero gsheet preglednico vpišemo
#   list .......... kateri list preglednice
#   znamka ........ znamka vozila
#   model ......... model vozila
#
preglednica="AvtoNet Yamaha motorji"
list="Podatki"
znamka="Yamaha"
model="TMAX"
kategorija="67000" # 61000=motorna kolesa, 67000 = maxi scooter.. drugega še ne poznam
oblika="6017" #6017 = MaxiScooter .. drugo je še treba dopolniti
ccm_min="500"
#
# ================================================================================================================================
# od tu dalje ne tikaj brez potrebe
#
require 'rubygems'
require 'nokogiri'
require 'open-uri'
require 'bundler'
Bundler.require

columnname=[]
columnname << "Slika"
columnname << "Zap.št."
columnname << "Id"
columnname << "Ogledov danes"
columnname << "Letnik"
columnname << "Cena"
columnname << "Prevoženi km"
columnname << "Km/Leto"
columnname << "Starost"
columnname << "Oglas oddan"
columnname << "1.registracija"
columnname << "Motor"
columnname << "Menjalnik"
columnname << "Zagon motorja"
columnname << "Prenos moči"
columnname << "Blokada"
columnname << "Kategorija"
columnname << "Barva"
columnname << "Interna številka"
columnname << "Kraj ogleda"
columnname << "Tehnični pregled"
columnname << "Rok dobave"
columnname << "Prodajalec"
columnname << "Telefon"

columnchar="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
#
# poveži ime s črko stolpca
#
#columnmap=[[],[]]
columnmap=[]
i=0
columnname.each do |c|
  columnmap << [c,columnchar[i]]
  i+=1
end
#
# Obdelava
#
puts "OglasiMotor | Berem avto.net"

#filter="https://www.avto.net/Ads/results.asp?znamka=#{znamka.gsub(" ","&")}&model=#{model.gsub(" ","&")}&modelID=&tip=&znamka2=&model2=&tip2=&znamka3=&model3=&tip3=&cenaMin=0&cenaMax=999999&letnikMin=0&letnikMax=2090&bencin=0&starost2=999&oblika=6004&ccmMin=0&ccmMax=99999&mocMin=&mocMax=&kmMin=0&kmMax=9999999&kwMin=0&kwMax=999&motortakt=0&motorvalji=0&lokacija=0&sirina=&dolzina=&dolzinaMIN=&dolzinaMAX=&nosilnostMIN=&nosilnostMAX=&lezisc=&presek=&premer=&col=&vijakov=&EToznaka=&vozilo=&airbag=&barva=&barvaint=&EQ1=1000000000&EQ2=1000000000&EQ3=1000000000&EQ4=100000000&EQ5=1000000000&EQ6=1000000000&EQ7=1110100120&EQ8=1010000006&EQ9=100000000&KAT=1060000000&PIA=&PIAzero=&PSLO=&akcija=&paketgarancije=&broker=&prikazkategorije=&kategorija=&zaloga=10&arhiv=&presort=&tipsort=&stran="
filter01="https://www.avto.net/Ads/results.asp?znamka=#{znamka.gsub(" ","&")}&model=#{model.gsub(" ","&")}&modelID=&"
filter02="tip=&znamka2=&model2=&tip2=&znamka3=&model3=&tip3=&cenaMin=0&cenaMax=999999&letnikMin=0&letnikMax=2090&bencin=0"
filter03="&starost2=999&oblika=#{oblika.to_s}&ccmMin=#{ccm_min}&ccmMax=99999&mocMin=&mocMax=&kmMin=0&kmMax=9999999&kwMin=0&kwMax=999"
filter04="&motortakt=0&motorvalji=0&lokacija=0&sirina=0&dolzina=&dolzinaMIN=0&dolzinaMAX=100&nosilnostMIN=0&nosilnost"
filter05="MAX=999999&lezisc=&presek=0&premer=0&col=0&vijakov=0&EToznaka=0&vozilo=&airbag=&barva=&barvaint=&"
filter06="EQ1=1000000000&EQ2=1000000000&EQ3=1000000000&EQ4=100000000&EQ5=1000000000&EQ6=1000000000&EQ7=1110100120"
filter07="&EQ8=1010000006&EQ9=1000000000&KAT=1060000000&PIA=&PIAzero=&PSLO=&akcija=0&paketgarancije=&broker=0"
filter08="&prikazkategorije=0&kategorija=#{kategorija.to_s}&zaloga=10&arhiv=0&presort=3&tipsort=DESC&stran=1&submodel=0"
filter="#{filter01}#{filter02}#{filter03}#{filter04}#{filter05}#{filter06}#{filter07}#{filter08}"
#puts filter
doc=Nokogiri::HTML(open(filter))
hrefs= doc.css("a").map {|element| element["href"]}
oglasi=[]
id_oglasov=[]
hrefs.each do |e|
  if e.include?"./park_mojavtonet.asp"
    #puts "https://www.avto.net/Ads/details.asp?id="+e[e.index('ID=')+3,8]
    oglasi << "https://www.avto.net/Ads/details.asp?id="+e[e.index('ID=')+3,8]
  end
end

#puts "OglasiMotor | izpis columnmap"
#columnmap.each do |n,v|
#  puts "#{n}='#{v}'"
#end
puts "OglasiMotor | povezujem gsheets"

session=GoogleDrive::Session.from_service_account_key("client_secret.json")
spreadsheet= session.spreadsheet_by_title(preglednica)
#worksheet= spreadsheet.worksheets[0]
#worksheet=spreadsheet.worksheets.first
#worksheet=spreadsheet.worksheet_by_title(list)
worksheet=spreadsheet.worksheet_by_title("Podatki")
puts "OglasiMotor | vpis vrednosti"
#
# glava
#
columnmap.each do |c|
  #puts "worksheet[\"#{c[1]}1\"]=#{c[0].to_s}"
  worksheet["#{c[1]}1"]=c[0].to_s
end
#
#   začnemo v vrstici 2
#
row=2

oglasi.each do |url_oglasa|
  #puts url_oglasa
  xml_oglasa=Nokogiri::HTML(open(url_oglasa))
  #
  # Slika
  #
  slika_url= xml_oglasa.xpath('//img[@class="OglasThumb"]').first["src"]
  attribute_name="Slika"
  attribute_value="=image(\"#{slika_url}\",1)"
  #puts "ATTRIBUTE=#{attribute_name} VALUE=#{attribute_value}"
  map=columnmap.select{|e| e[0] == attribute_name}
  worksheet["#{map[0][1]}#{row}"]= attribute_value
  #
  # Zap. št. oglasa
  #
  map=columnmap.select{|e| e[0] == "Zap.št."}
  worksheet["#{map[0][1]}#{row}"]=row-1
  #
  # Id oglasa
  #
  id=url_oglasa[url_oglasa.index('id=')+3,8]
  #puts "ATTRIBUTE=ID VALUE=#{id}"
  map=columnmap.select{|e| e[0] == "Id"}
  #worksheet["#{map[0][1]}#{row}"]=id
  worksheet["#{map[0][1]}#{row}"]="=HYPERLINK(\"#{url_oglasa}\",\"#{id}\")"
  #worksheet.save
  #
  # Cena
  #
  price=xml_oglasa.css('div').xpath('div[@class="OglasPrice"]').text.delete(' .€').strip
  #puts "ATTRIBUTE=price VALUE=#{price}"
  map=columnmap.select{|e| e[0] == "Cena"}
  worksheet["#{map[0][1]}#{row}"]=price
  #
  # Oglas oddan
  #
  oglasoddan=xml_oglasa.css('div').xpath('div[@class="OglasContactStatLeftRecap"]').text.strip
  if oglasoddan!=nil
    #puts "ATTRIBUTE=Oglas oddan VALUE=#{oglasoddan[13,20]}"
    attribute_value=oglasoddan[13,20]
  else
    #puts "ATTRIBUTE=Oglas oddan VALUE="
    attribute_value=""
  end
  if attribute_value==nil
    attribute_value=""
  end
  map=columnmap.select{|e| e[0] == "Oglas oddan"}
  worksheet["#{map[0][1]}#{row}"]=attribute_value
  #worksheet.save
  #
  # Ogledov danes
  #
  ogledovdanes=xml_oglasa.css('div').xpath('div[@class="OglasContactStatRightRecap"]').text.strip
  if ogledovdanes!=nil
    #puts "ATTRIBUTE=Ogledov danes VALUE=#{ogledovdanes[17,10]}"
    attribute_value=ogledovdanes[17,10]
    if attribute_value==nil
      attribute_value="0"
    end
  else
    #puts "ATTRIBUTE=Ogledov danes VALUE=neznano"
    attribute_value="0"
  end
  #puts "ATTRIBUTE=Ogledov danes VALUE=#{attribute_value}"
  #puts "#{map[0][1]}#{row}=#{attribute_value}"
  map=columnmap.select{|e| e[0] == "Ogledov danes"}
  worksheet["#{map[0][1]}#{row}"]=attribute_value
  #worksheet.save
  #
  # Prodajalec
  #
  prodajalec=xml_oglasa.css('div').xpath('div[@class="PaddingTopBtm"]').text.strip
  if prodajalec!=nil
    #puts "ATTRIBUTE=Ogledov danes VALUE=#{ogledovdanes[17,10]}"
    attribute_value=prodajalec
    if attribute_value==nil
      attribute_value=""
    end
  else
    #puts "ATTRIBUTE=Prodajalec VALUE=neznano"
    attribute_value=""
  end
  map=columnmap.select{|e| e[0] == "Prodajalec"}
  #puts "Prodajalec=#{attribute_value}"
  worksheet["#{map[0][1]}#{row}"]=attribute_value
  #
  # Telefon
  #
  telefon=xml_oglasa.css('div').xpath('div[@class="OglasMenuBox Bold OglasMenuBoxPhone"]').text.strip
  if telefon!=nil
    #puts "ATTRIBUTE=Ogledov danes VALUE=#{ogledovdanes[17,10]}"
    attribute_value="#{telefon}"
    if attribute_value==nil
      attribute_value=""
    end
  else
    #puts "ATTRIBUTE=Prodajalec VALUE=neznano"
    attribute_value=""
  end
  map=columnmap.select{|e| e[0] == "Telefon"}
  #puts "Telefon=#{attribute_value}"
  worksheet["#{map[0][1]}#{row}"]=attribute_value
  #
  # ostali stolpci Letnik, Starost, Prevoženih km, ...
  #
  letnik=nil
  km=nil
  xml_oglasa.css('div').xpath('div[@class="OglasData"]').each do |e|
      attribute_name=e.xpath('div[@class="OglasDataLeft"]').text.delete(':')
      attribute_value=e.xpath('div[@class="OglasDataRight"]').text.strip
      if attribute_name!=nil and attribute_name!=""
        #puts "ATTRIBUTE=#{attribute_name} | VALUE=#{attribute_value}"
        map=columnmap.select{|e| e[0] == attribute_name }
        if "#{map[0][1]}#{row}"!=nil
          if attribute_value==nil
              attribute_value=""
          end
          worksheet["#{map[0][1]}#{row}"]=attribute_value
          #worksheet.save
        end
      else
        #puts "*Attribute name or value is nil* ATTRIBUTE=#{attribute_name} VALUE=#{attribute_value}"
      end
      if attribute_name=="Letnik"
          letnik=attribute_value.to_i
      end
      if attribute_name=="Prevoženi km"
          km=attribute_value.to_i
      end
      if letnik!=nil and km!=nil
        if letnik!=Time.now.year
            # zapiši
            km_na_leto=km.to_f/(Time.now.year.to_f-letnik.to_f)
            attribute_name="Km/Leto"
            attribute_value="#{km_na_leto.to_i}"
            map=columnmap.select{|e| e[0] == attribute_name }
            #puts "Km/Leto=#{attribute_value.to_i}"
            worksheet["#{map[0][1]}#{row}"]=attribute_value
        end
      end
  end
  row+=1
  #worksheet.save
end
puts "OglasiMotor | shrani v gsheets"
worksheet.save



