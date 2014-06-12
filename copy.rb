# encoding: utf-8
require 'axlsx' # je samo za pisanje 
require 'spreadsheet' # ovo koristim za otvaranje
require 'active_support/core_ext' # ovo je da bi mi uspeo downcase na Ć, i deo njega je onaj mb_char

#USLOV #1 - excel fajl mora biti u istom direktorijumu
#USLOV #2 - celiju "ime" drzati na mestu [1,1] - najbolje
#USLOV #3 - ne vrsim upis u vec postojeci fajl, nego kreiram novi

#funkcija vraca Array sa jednim clanom, cije su x i y vrednosti 
#koordinate celije sa sadrzajem koju trazimo
def broj_celije(ime="ime")
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet.open('PocetniSample.xls')
  sheet = book.worksheet('prvi')
  uslov = false
  
  until uslov == true
  
  x = 0
 # each row ne uzima u obzir prazne redove uopste, pa sam dodao bla u excelu
  sheet.each do |row|
   y = 0 
   x += 1
    row.each do |cell|
      y += 1
        if cell.to_s == ime
          uslov = true
          #return ("Broj reda je: #{x}, broj kolone(zero based) je: #{y}")
          return x-1, y-1 
        #else
         # puts "Nema takve kolone"
        end
      end
    end
  end
end

# funkcija vraca sadrzaj celije cije koordinate ubacujemo
# kao array sa jednim clanom
def pronadji_po_broju(x)
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet.open('PocetniSample.xls')
  sheet = book.worksheet('prvi')
  #zato sto su rowovi 0-based
  a = sheet.row(x[1])
  #zato sto je array - a on je 0-based
  b = a[x[0]].to_s
  return b
end

# funkcija vraca Array koji sadrzi sve redove ispod koordinata koji joj ubacimo
# + jos dve kolone udesno
def printrows_to_array(x)
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet.open('PocetniSample.xls')
  sheet = book.worksheet('prvi')
  array = []
  hash = {}
  #ne moram da koristim sheet.drop(x[0]).each do |row| - jer podrazumevam da iznad nemam nista
  sheet.drop(x[0]+1).each do |row|#ovde je samo +1, da ne bih dobio "ime" u arrayu - sada isljucio
  array << [row[x[1]].to_s, row[x[1]+1].to_s, row[x[1]+2].to_s] #ovo je multidimensional
  end
  return array
end

#idemo dalje, taj array obacujemo u neki metod - mozda posle probati i sa otvaranjem ARRAY klase?

#funcija koja pretrazuje O(n2) slozenosti, brise duplikate po string_isti kriterijumu
# i vraca array koji sadrzi samo jedinstvena imena, sa zbirovima vrednosti u kolonama
# 2 i 3 
def nadji_duplikate(x)
  upamtiGdeJePrviPrimerak = 0
  #moracu preko 2 arraya
  b = Array.new(x)
  x.each do |ime|
    brojac = 0
    brojponavljanja = 0
    while brojac < b.length do

      if string_isti(ime[0], b[brojac][0],90) == true
          brojponavljanja += 1
          if (brojponavljanja == 2 or brojponavljanja == 3)
            b[brojac][0] = "prazno"
            b[upamtiGdeJePrviPrimerak][1] = b[upamtiGdeJePrviPrimerak][1].to_f + b[brojac][1].to_f
            b[upamtiGdeJePrviPrimerak][2] = b[upamtiGdeJePrviPrimerak][2].to_f + b[brojac][2].to_f
            brojac += 1
          elsif brojponavljanja == 1 
            upamtiGdeJePrviPrimerak = brojac
          else
            brojac += 1 
          end
          brojac += 1
      else
        brojac += 1
      end
    
    end
    #ovde da ubacim brisanje praznih
    b.delete_if {|x| x[0] == "prazno"}
  end
   return b  
end

# funkcija koja mi sluzi kao kriterijum za jednakost poverioca-stringova
def string_isti(a,b, c = 100)
  x = 0
  y = 0
  i = 0
  a_lower = a.mb_chars.downcase.to_s #probati sa .downcase, i videce se da se Š i Ć Č ne downcaseu-ju
  b_lower = b.mb_chars.downcase.to_s # sa 300 sekundi sa ovim skocilo na 800 :(
  
  while x < a.length
      if a_lower[x] == "š"
         a_lower[x] = "s"
         x += 1
      elsif
         a_lower[x] == "ć"
         a_lower[x] = "c"
         x += 1
      elsif
         a_lower[x] == "č"
         a_lower[x] = "c"
         x += 1
      else
        x += 1
       end
  
  end
  x = 0
  while x < b.length
      if b_lower[x] == "š"
       b_lower[x] = "s"
       x += 1
      elsif
        b_lower[x] == "ć"
        b_lower[x] = "c"
        x += 1
      elsif
        b_lower[x] == "č"
        b_lower[x] = "c"
        x += 1
      else
        x += 1
      end
  end
  x = 0

  while i < a_lower.length # ne sme <= jer ce vrteti jedanput vise, a ruby nema out of range
    if a_lower[i] == b_lower[i]
      #puts a_lower[i] # ovo ce mi odstampati celu rec, i stace tamo gde se prvo slove ne slozi
      #puts a_lower.length
      y += 1
      i += 1
    else
      #return false # ovo smanjiju vreme za 1 minut - mislio sam da ce biti vise
      # ovo gore moram odcekirati ako zelim da koristim manje od 100%
      i += 1
    end
  end
  
  d = a_lower.length
  
  if y.to_f/d*100 >= c.to_f
    return true
  else
    return false
  end
  
end

# funkcija kojom vrsim upis arraya sa jedinstvenim vrednostima u xlsx fajl
def upisi_nove_redove(x)
  p = Axlsx::Package.new
  wb = p.workbook
  
  wb.add_worksheet(:name => "PrviSredjeno") do |sheet|
    x.each do |red|
    sheet.add_row [red[0], red[1], red[2]]
  end

end
  p.use_shared_strings = true #ne znam sta je
  p.serialize('KrajnjiSample.xlsx')
end

x = Time.now

#puts broj_celije()

@x = printrows_to_array(broj_celije())
puts "Broj redova u prvom arrayu je: #{@x.length}"
# @x je array sa svim podacima
#print @x

@y = nadji_duplikate(@x)
#@y je array sa jedinstvenim podacima
#puts "__________"
puts "Broj redova nakon sredjivanja je: #{@y.length}"
#puts @y

upisi_nove_redove(@y)
#upisujemo @y u excel fajl

y = Time.now
z = y - x 
puts "Utroseno vreme: #{z} sekundi"

