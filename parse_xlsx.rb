#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'roo'

def department_and_sickbed(sheet, row_num)
  departments = []
  sickbed = {}

  index = 0
  loop do
    tmp_department = sheet.cell(row_num + index, 'I')
    break if tmp_department.nil?

    tmp_department_array = tmp_department.split
    if tmp_department_array.size == 2 && tmp_department_array.last.match(NUMBER_REGEXP)
      sickbed[tmp_department_array.first] = tmp_department_array.last
    else
      departments << tmp_department
    end
    index += 1
  end

  [departments.join(' '), sickbed]
end

def number_of_each_health_care_worker(sheet, row_num)
  fulltime_doctor_num = 0
  fulltime_pharmacist_num = 0
  fulltime_dentist_num = 0
  parttime_doctor_num = 0
  parttime_pharmacist_num = 0
  parttime_dentist_num = 0

  index = 1
  loop do
    worker = sheet.cell(row_num + index, 'E')
    break if worker.nil?

    if worker.match(/常　勤/)
      index2 = 1
      loop do
        worker = sheet.cell(row_num + index + index2, 'E')
        break if worker.nil? || worker.match(/非常勤/)
        if worker.match(/\(医/)
          fulltime_doctor_num = worker.split.last.sub(')', '')
        elsif worker.match(/\(薬/)
          fulltime_pharmacist_num = worker.split.last.sub(')', '')
        end
        index2 += 1
      end
    elsif worker.match(/非常勤/)
      index2 = 1
      loop do
        worker = sheet.cell(row_num + index + index2, 'E')
        break if worker.nil?
        if worker.match(/\(医/)
          parttime_doctor_num = worker.split.last.sub(')', '')
        elsif worker.match(/\(薬/)
          parttime_pharmacist_num = worker.split.last.sub(')', '')
        end
        index2 += 1
      end
    end

    index += 1
  end

  [
    fulltime_doctor_num,
    fulltime_pharmacist_num,
    fulltime_dentist_num,
    parttime_doctor_num,
    parttime_pharmacist_num,
    parttime_dentist_num
  ]
end

def parse_medical_institution(sheet)
  lines = []

  1.upto(sheet.last_row) do |row_num|
    if sheet.cell(row_num, 'A')&.match(NUMBER_REGEXP)
      num, institution_num, institution_name, address, tel, founder, manager, registered_on, _, note = sheet.row(row_num)
      registered_reason = sheet.cell(row_num + 1, 'H')
      register_started_on = sheet.cell(row_num + 2, 'H')

      department, sickbed = department_and_sickbed(sheet, row_num)
      fulltime_doctor_num, fulltime_pharmacist_num, fulltime_dentist_num, parttime_doctor_num, parttime_pharmacist_num, parttime_dentist_num = number_of_each_health_care_worker(sheet, row_num)

      lines << [
        num,
        institution_num,
        institution_name,
        address,
        tel,
        founder,
        manager,
        registered_on,
        registered_reason,
        register_started_on,
        department,
        sickbed.map { |k, v| "#{k}(#{v})" }.join(' '),
        fulltime_doctor_num.to_i + fulltime_pharmacist_num.to_i + fulltime_dentist_num.to_i,
        fulltime_doctor_num,
        fulltime_pharmacist_num,
        fulltime_dentist_num,
        parttime_doctor_num.to_i + parttime_pharmacist_num.to_i + parttime_dentist_num.to_i,
        parttime_doctor_num,
        parttime_pharmacist_num,
        parttime_dentist_num,
        note,
      ].join(SEPARATOR)
    end
  end

  lines
end

filename = 'kumamoto_ika_02.xlsx'
xlsx = Roo::Excelx.new(filename)
# puts xlsx.info
sheet = xlsx.sheet(0)

NUMBER_REGEXP = /\A[1-9]\d*\z/
SEPARATOR = ","
HEADER = %w(
  項番
  医療機関番号
  医療機関名称
  医療機関所在地
  電話番号
  開設者氏名
  管理者氏名
  指定年月日
  登録理由
  指定期間始
  診療科名
  病床数
  常勤
  常勤(医)
  常勤(薬)
  常勤(歯)
  非常勤
  非常勤(医)
  非常勤(薬)
  常勤勤(歯)
  備考
)

puts HEADER.join(SEPARATOR)
puts parse_medical_institution(sheet)