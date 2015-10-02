% Copyright by D. Walczak
% Rev 0
% 2015/09
% This short function-script takes as an input two arguments:
% 1) some UK city as string value
% 2) maximum price as double value
% and then goes to the website: http://www.rightmove.co.uk, performs
% automatically the set serach and after a while returns an EXCEL file list 
% of flats offered for sale in that city up to that price inclusive.
% city = 'Liverpool', 'Glasgow', etc.
% max_price = 40000, 50000, 100000, 150000, 200000, 250000, ...
% EXCEL file, stored in current working directory is of a format:
% flats_prices_city_current_date.xlsx


function ans = rightmove_properties(city, max_price) 

% Determine the number of subpages returned from search result
A = urlread(['http://www.rightmove.co.uk/property-for-sale/' city '.html?sortType=6&maxPrice=' num2str(max_price) '&displayPropertyType=flats&numberOfPropertiesPerPage=50']);
i = regexp(A, '>\d+</a></li>', 'match');
i = regexprep(i, '>|(</a></li>)', '');
if isempty(i)
    i = {'0'};
else
    i = [{'0'}, i];
end
i = 50*(char(i) - 48)';
A = [];
for j = 1:numel(i)
    A = [A; {urlread(['http://www.rightmove.co.uk/property-for-sale/' city '.html?sortType=6&maxPrice=' num2str(max_price) '&displayPropertyType=flats&numberOfPropertiesPerPage=50&index=' num2str(i(j))])}];
end

A = [A{:}];
I = regexp(A, '<p class="price">(.*?)<p>', 'match');

% extract property ID
PropertyID = regexp(I, '"[0-9]{8}"', 'match');
PropertyID = strrep([PropertyID{:}], '"', '');

% extract price
PricesBelow100K = regexp(I, '&pound;([0-9][0-9],[0-9][0-9][0-9])</sp', 'match');
PricesBelow100K = regexprep([PricesBelow100K{:}],'(</sp)|(,)|(&pound;)', '');

% extract property address descriptions
Addresses = regexp(I, 'address">(.*?)</sp', 'match');
Addresses = regexprep([Addresses{:}], '(address">)|(</sp)|(\n)|(\t)|(<strong>)|(</strong>)', '');

% extract agent name
AgentName = regexp(I, '(Added|Reduced)(.*?)by(.*?)<span class', 'match');
AgentName = regexprep([AgentName{:}], 'Added|Reduced ((today|yesterday)</span> )|(on [0-9]{2}\/[0-9]{2}\/[0-9]{4}) |by|<span class|<b>|</b>|\t|\n','');

% extract agent phone contact
PhoneNr = regexp(I, 'strong>Call: (.*?)</str', 'match');
PhoneNr = regexprep([PhoneNr{:}], '(strong>Call: )|(</str)', '');

% Write data to EXCEL file
headers = {'Property ID', 'Price GBP', 'Address', 'Agency Name', 'Agency Contact Nr'};
filename = [pwd '\flats_prices_' [city] '_' [date] '.xlsx'];
xlswrite(filename, [headers; PropertyID' PricesBelow100K' Addresses' AgentName' PhoneNr']);
hExcel = actxserver('Excel.Application');
hWorkbook = hExcel.Workbooks.Open(filename);
hWorksheet = hWorkbook.Sheets.Item(1);
hWorkbook.Save
hWorkbook.Close
hExcel.Quit

% success announcment
fprintf('\nDone!! Found and printed %d records to:\n%s.\n', numel(I), filename);
    
end
