<?php
namespace Vinhnt\Databasedocs\Providers;

use Illuminate\Support\ServiceProvider;

class DatabaseDocsServiceProvider extends ServiceProvider {
    public function boot()
    {
        $this->loadRoutesFrom(__DIR__.'/../routes/routes.php');
    }
    public function register()
    {

    }
}