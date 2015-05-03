angular.module('myApp', ['ngRoute'])
.constant('routeConstants',
	{STUDENT_BASE_PREFIX:'/students/'})
.config(['$routeProvider',
  function($routeProvider) {
    $routeProvider
    .when('/studentDetails/:studentId', {
        templateUrl: 'templates/student-details.tmpl.html',
        controller: 'StudentDetailsController',
        resolve:{"initData":studentDetailsController.init
				}
    })
    .when('/',
	{
		templateUrl: 'templates/landing.tmpl.html',
        controller: 'StudentController'
       
	})
.otherwise({
	redirectTo: '/'
      });
}]);;
