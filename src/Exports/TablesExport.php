<?php
namespace Vinhnt\Databasedocs\Exports;

use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Illuminate\Support\Facades\DB;
use Vinhnt\Databasedocs\Exports\Sheets\TableSheet;
use Vinhnt\Databasedocs\Exports\Sheets\CoverSheet;
use Vinhnt\Databasedocs\Exports\Sheets\ERDSheet;
use Vinhnt\Databasedocs\Exports\Sheets\ListOfTablesSheet;
use Vinhnt\Databasedocs\Exports\Sheets\TableOfContentsSheet;
use Vinhnt\Databasedocs\Exports\Sheets\UpdateHistorySheet;

class TablesExport implements WithMultipleSheets
{
    use Exportable;

    /**
     * @return array
     */
    public function sheets(): array
    {
        $sheets = [];

        $sheets[] = new CoverSheet();
        $sheets[] = new UpdateHistorySheet();
        $sheets[] = new TableOfContentsSheet();
        $sheets[] = new ERDSheet();

        $tableName = $this->getTableNameAndDescription()[0];
        $tableDescription = $this->getTableNameAndDescription()[1];
        $sheets[] = new ListOfTablesSheet($tableName, $tableDescription);

        $tables_array = $this->tablesArray();
        foreach ($tables_array as $table) {
            $sheets[] = new TableSheet($table['table_name'], $table['fields'], $table['fields_info']);
        }
        
        return $sheets;
    }

    /**
     * @return array
     */
    public function getTableNameAndDescription()
    {
        $names = [];
        $description = [];

        $database = env('DB_DATABASE');

        $tables = DB::select("SELECT table_name, table_comment FROM information_schema.tables WHERE table_schema = '$database'");
        foreach ($tables as $table) {
            $names[] = $table->TABLE_NAME;
            $description[] = $table->TABLE_COMMENT;
        }

        return [$names, $description];
    }

    /**
     * @return array
     */
    public function tablesArray() {
        $tables = DB::getDoctrineSchemaManager()->listTables();

        $tables_array = [];
        foreach ($tables as $table) {
            $table_name = $table->getName();

            $table_fields = [];
            $table_fields_info = [];

            foreach ($table->getColumns() as $column) {
                $name = $column->getName();
                $type = $column->getType()->getName();
                $notnull = $column->getNotnull();
                $autoIncrement = $column->getAutoincrement();
                $default = $column->getDefault();
                $length = $column->getLength();
                $description = $column->getComment();
                $primaryKey = $table->getPrimaryKey() ? $table->getPrimaryKey()->getColumns()[0] === 'id' : false;
                $uniqueKey = false;
                $foreignKey = false;
            
                // Check if this column is a unique key
                foreach ($table->getIndexes() as $index) {
                    if ($index->getColumns()[0] === $name) {
                        $uniqueKey = $index->isUnique();
                        break;
                    }
                }

                // Check if this column is a foreign key
                foreach ($table->getForeignKeys() as $fK) {
                    if (in_array($name, $fK->getLocalColumns())) {
                        $foreignKey = true;
                        break;
                    }
                }

                $table_fields[] = $name;

                $table_fields_info[] = [
                    'primary_key' => $primaryKey,
                    'unique' => $uniqueKey,
                    'notnull' => $notnull,
                    'auto_increment' => $autoIncrement,
                    'foreign_key' => $foreignKey,
                    'type' => $type,
                    'length' => $length,
                    'default' => $default,
                    'description' => $description,
                ];
            }

            array_push($tables_array, [
                'table_name' => $table_name,
                'fields' => $table_fields,
                'fields_info' => $table_fields_info,
            ]);
        }

        return $tables_array;
    }
}