(function (angular) {
    angular.module('wcjc', []);

    angular.module('wcjc').controller('indexController', indexController);

    indexController.$inject = ["$scope"];

    function indexController($scope) {
        var vm = this;
        vm.Attendees = [];
        vm.AttendeesArray = [];

        vm.handleDrop = handleDrop;
        vm.handleFile = handleFile;

        function handleDrop(e) {
            e.stopPropagation();
            e.preventDefault();
            var files = e.dataTransfer.files;
            var i,f;
            for (i = 0, f = files[i]; i != files.length; ++i) {
                var reader = new FileReader();
                var name = f.name;
                reader.onload = function(e) {
                    var data = e.target.result;

                    /* if binary string, read with type 'binary' */
                    var workbook = XLSX.read(data, {type: 'binary'});

                    /* DO SOMETHING WITH workbook HERE */

                    var sheet_name_list = workbook.SheetNames;
                    sheet_name_list.forEach(function(y) { /* iterate through sheets */
                        var worksheet = workbook.Sheets[y];
                        for (z in worksheet) {
                            /* all keys that do not begin with "!" correspond to cell addresses */
                            if(z[0] === '!') continue;
                            console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
                        }
                    });
                };
                reader.readAsBinaryString(f);
            }
        }

        function handleFile(e) {
            var files = e.target.files;
            var i,f;
            for (i = 0, f = files[i]; i != files.length; ++i) {
                var reader = new FileReader();
                var name = f.name;
                reader.onload = function(e) {
                    var data = e.target.result;

                    var workbook = XLSX.read(data, {type: 'binary'});

                    /* DO SOMETHING WITH workbook HERE */
                    var sheet_name_list = workbook.SheetNames;
                    sheet_name_list.forEach(function(y) { /* iterate through sheets */
                        var worksheet = workbook.Sheets[y];

                        var rowNumber = 0;
                        var currentRowNumber = 0;
                        var currentObject = {};

                        for (z in worksheet) {

                            currentRowNumber = ~~z.substring(1);

                            if(rowNumber != currentRowNumber)
                            {
                                if(rowNumber > 1)
                                {
                                    vm.Attendees.push(angular.copy(currentObject));
                                }

                                currentObject = {};
                                rowNumber = currentRowNumber;
                            }

                            /* all keys that do not begin with "!" correspond to cell addresses */
                            if(z[0] === '!') continue;
                            //console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));

                            if(z.indexOf('A') > -1)
                            {
                                currentObject.Name = worksheet[z].v;
                            }
                            else if(z.indexOf('B') > -1) {
                                currentObject.LastName = worksheet[z].v;
                            }
                            else if(z.indexOf('E') > -1)
                            {
                                currentObject.Club = worksheet[z].v;
                            }
                            else if(z.indexOf('G') > -1)
                            {
                                currentObject.SleepOver = worksheet[z].v != 'Nej';
                            }

                        }

                        vm.Attendees.push(angular.copy(currentObject));
                        $scope.$digest();
                    });
                };
                reader.readAsBinaryString(f);
            }
        }

        var uploadElement = document.getElementById("xlf");
        uploadElement.addEventListener('change', vm.handleFile, false);
        //dropElement.addEventListener('drop', handleDrop, false);
    }
}(window.angular));
