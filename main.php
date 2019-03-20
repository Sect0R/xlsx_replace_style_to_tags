<?php
class main
{
    /**
     * Main function for normalize text
     * @param string $text
     * @param \PhpOffice\PhpSpreadsheet\Style\Style|\PhpOffice\PhpSpreadsheet\RichText\ITextElement $style
     * @return bool
     */
    public static function replaceStyleToTags(&$text, $style)
    {
        if (!$text) {
            return false;
        }

        $trimmedText = trim($text);

        if (!$trimmedText) {
            return false;
        }

        $elementUpdates = [];

        // add styled tags (<strong><u><i>)
        $elementUpdates[] = self::addTagsByFontStyle($style->getFont(), $text);

        // replace nl to br
        $elementUpdates[] = self::addBrTag($text);

        return array_search(true, $elementUpdates) !== false;
    }

    /**
     * Add tags to the styled text or cells
     * @param \PhpOffice\PhpSpreadsheet\Style\Font $font
     * @param string $text
     * @return bool Element updated
     */
    public static function addTagsByFontStyle($font = null, &$text = '')
    {
        if (!$font) {
            return false;
        }

        $tagsAdded = [];

        // replace bold to <strong>
        if ($font->getBold() === true) {
            $tagsAdded[] = self::addTagToText($text, '<strong>');
        }

        // $font->getUnderline() can be single or double
        if ($font->getUnderline() && $font->getUnderline() != 'none') {
            $tagsAdded[] = self::addTagToText($text, '<u>');
        }

        if ($font->getItalic() === true) {
            $tagsAdded[] = self::addTagToText($text, '<i>');
        }

        return array_search(true, $tagsAdded) !== false;
    }

    /**
     * Add tag to text (example: '<strong>Text</strong>'
     * @param string $text Text
     * @param string $tag Tag with <>, example <strong>
     * @return bool
     */
    public static function addTagToText(&$text, $tag)
    {
        if (!preg_match('/^' . $tag . '(.+?)' . '<\/' . substr($tag, 1) . '$/is', $text)) {
            $prefix = '';
            $postfix = '';

            if ($text[0] == ' ') {
                $prefix = ' ';
                $text = ltrim($text);
            }

            if ($text[strlen($text) - 1] == ' ') {
                $postfix = ' ';
                $text = rtrim($text);
            }

            $text = $prefix . $tag . $text . '</' . substr($tag, 1) . $postfix;

            return true;
        }

        return false;
    }

    /**
     * Replace \n or \r\n to <br />
     * @param string $text
     * @return bool
     */
    public static function addBrTag(&$text = '')
    {
        $findBrRegexp = '/\v+|\\\r\\\n/Ui';

        if (preg_match($findBrRegexp, $text) && strpos($text, '<br />') === false) {
            $text = preg_replace($findBrRegexp, '<br/>' . "\n", $text);
            return true;
        }

        return false;
    }
}