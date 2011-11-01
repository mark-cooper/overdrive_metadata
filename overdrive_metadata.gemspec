# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %q{overdrive_metadata}
  s.version = "1.0.0.20111101103802"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Mark Cooper"]
  s.date = %q{2011-11-01}
  s.description = %q{FIX (describe your package)}
  s.email = ["markchristophercooper@gmail.com"]
  s.extra_rdoc_files = ["History.txt", "Manifest.txt", "README.txt"]
  s.files = [".autotest", "History.txt", "Manifest.txt", "README.txt", "Rakefile", "lib/overdrive_metadata.rb", "raw/test.xls", "test/test_overdrive_metadata.rb", ".gemtest"]
  s.homepage = %q{http://www.librodoor.net}
  s.rdoc_options = ["--main", "README.txt"]
  s.require_paths = ["lib"]
  s.rubyforge_project = %q{overdrive_metadata}
  s.rubygems_version = %q{1.5.2}
  s.summary = %q{FIX (describe your package)}
  s.test_files = ["test/test_overdrive_metadata.rb"]

  if s.respond_to? :specification_version then
    s.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_development_dependency(%q<hoe>, ["~> 2.12"])
    else
      s.add_dependency(%q<hoe>, ["~> 2.12"])
    end
  else
    s.add_dependency(%q<hoe>, ["~> 2.12"])
  end
end
