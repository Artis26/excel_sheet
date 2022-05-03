<?php
namespace App\Models;

class Person {
    private string $name;
    private string $position;

    public function __construct(string $name, string $position) {
        $this->name = $name;
        $this->position = $position;
    }

    public function getName(): string {
        return $this->name;
    }

    public function getPosition(): string {
        return $this->position;
    }
}
