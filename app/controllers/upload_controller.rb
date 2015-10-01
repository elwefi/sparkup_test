class UploadController < ApplicationController

  def new

  end

  def create
    spreadsheet = open_spreadsheet(params[:file])
    @final_contacts = []
    @fail_contacts = []

    (2..spreadsheet.last_row).each do |i|
      line = spreadsheet.row(i)

      # Vérifier si même nom et prénom ou même adresse email
      if @final_contacts.detect{|contact| (contact[:first_name] == line[0].capitalize && contact[:last_name] == line[1].capitalize) || contact[:email] == line[2]}.blank?
        
        # Vérifier adresse email
        if (line[2]=~ /\A([^@\s]+)@((?:[-a-z0-9]+\.)+[a-z]{2,})\Z/) == nil
          @fail_contacts << {first_name: line[0], last_name: line[1], email: line[2],reason: "Email Invalide" }
        # Vérifier si moins de 3 lettres dans le prénom
        elsif line[0].length < 3
          @fail_contacts << {first_name: line[0], last_name: line[1], email: line[2],reason: "Prénom Invalide" }
        # vérifier si moins de 3 lettres dans le nom
        elsif line[1].length < 3
          @fail_contacts << {first_name: line[0], last_name: line[1], email: line[2],reason: "Nom Invalide" }
        # remove caractères invalides et mettre en majuscule le Nom et le Prénom
        else 
          @final_contacts << {first_name: line[0].gsub(/(\W|\d)/, "").capitalize, last_name: line[1].gsub(/(\W|\d)/, "").capitalize, email: line[2]}
        end

      else
        @fail_contacts << {first_name: line[0], last_name: line[1], email: line[2],reason: "Fusion" }
      end
    end

  end

  def open_spreadsheet(file)
    case File.extname(file.path)
    when ".xls" then Roo::Excel.new(file.path)
    when ".xlsx" then Roo::Excelx.new(file.path)
    else raise "Unknown file type: "
    end
  end

end
