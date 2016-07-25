Hyper
=====

Hyper lets you automatically populate Google Docs with data from arbitrary URLs. It comes with GitHub authorization so that can use URLs pointing to content hosted on GitHub. A great initial use case is Jupyter notebooks on GitHub.

Installation
------------

The easiest way to install `Hyper` is [here](https://chrome.google.com/webstore/detail/hyper/lbiblkbfmhelinhcobciimdmcnngcffp?utm_source=permalink) via the Google Docs add-on store. Please reply [here](https://github.com/tavinathanson/hyper/issues/1) to be added as a private tester so that you can install `Hyper` this way.

Usage
-----

* Choose a URL as a source of data. For example, a Jupyter notebook.
* In that source of data, use the format `{{{<label>:<insert arbitrary text>}}}`. For example, `{{{median_mutations:234}}}`.
* Create a hyperlink in your Google Doc with any text and the following url: `<your source of data URL>?hyper=<label>`. For example, `https://github.com/myorg/myrepo/tree/master/analyses/notebooks/notebook.ipynb?hyper=median_mutations`.
* Go to `Add-ons -> Hyper -> Hyperize Links` in your Google Doc, and watch the text of your hyperlink transform into the data. For example, into `234`.
* When the data updates, feel free to Hyperize links as needed.

Icons made by [SimpleIcon](http://www.flaticon.com/authors/simpleicon) from [Flaticon](http://www.flaticon.com), licensed by [CC 3.0 BY](http://creativecommons.org/licenses/by/3.0/).
