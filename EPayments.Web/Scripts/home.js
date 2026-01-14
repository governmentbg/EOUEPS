$(document).ready(function () {
    $("button.btn-tabs-ham").click(function () {
        $("#nav-login-reg").toggleClass("collapsed")
    });
    $("button.btn-close").click(function () {
        $(this).closest(".wrapper").addClass("hidden")
    });
    $(".btn_home_entrance").click(function () {
        $(this).siblings("li").removeClass("active");
        $(this).addClass("active");
        $("#nav-login-reg").addClass("collapsed");
        $("#entrance").siblings(".home-tabs-target").removeClass("show");
        $("#entrance").addClass("show")
    });
    $(".btn_home_eauproviders").click(function () {
        $(this).siblings("li").removeClass("active");
        $(this).addClass("active");
        $("#nav-login-reg").addClass("collapsed");
        $("#eauproviders").siblings(".home-tabs-target").removeClass("show");
        $("#eauproviders").addClass("show")
    });
    $(".btn_home_serviceproviders").click(function () {
        $(this).siblings("li").removeClass("active");
        $(this).addClass("active");
        $("#nav-login-reg").addClass("collapsed");
        $("#serviceproviders").siblings(".home-tabs-target").removeClass("show");
        $("#serviceproviders").addClass("show")
    });
    $(".btn-select-profile").click(function () {
        $(this).toggleClass("selected");
        $(".nav-select-profile").toggleClass("selected")
    });
    $(".btn-left-nav-menu").click(function () {
        $(".left-nav-menu").toggleClass("show");
        $(".left-nav-wrapper").toggleClass("collapsed")
    });
    $(".nav-select-profile li").click(function () {
        $(".btn-select-profile").removeClass("selected");
        $(this).toggleClass("selected");
        $(this).siblings("li").removeClass("selected");
        $(".nav-select-profile").toggleClass("selected")
    });
    $(".tabs-nav button").click(function () {
        $(this).addClass("active");
        $(this).siblings("button").removeClass("active")
    });
    $(".btn-search-filter").click(function () {
        $(this).closest(".search-filter").toggleClass("show")
    });

    $(".seemore").click(function () {
        $(".dots").toggle();
        $(".more").toggle();
        if ($(this).text() == "Научи повече") {
            $(this).text("Затвори");
        } else {
            $(this).text("Научи повече");
        }
    });
    $('.selectpicker').selectpicker({
        liveSearch: true,
        showSubtext: true
    });
    $(function () {
        $(".chosen").chosen();
    });
    $(function () {
        $(".chosenobligation").disable;
    });
})